#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <structmember.h>
#include <datetime.h>
#include <math.h>

/* ===== Types cache (imported from Python) ===== */
static PyObject *decimal_type = NULL;
static PyObject *datetime_type = NULL;
static PyObject *date_type = NULL;
static PyObject *timedelta_type = NULL;

static int _ensure_types(void)
{
    if (decimal_type && datetime_type && date_type && timedelta_type)
        return 1;

    if (!decimal_type)
    {
        PyObject *mod = PyImport_ImportModule("decimal");
        if (!mod) { PyErr_Clear(); return 0; }
        decimal_type = PyObject_GetAttrString(mod, "Decimal");
        Py_DECREF(mod);
        if (!decimal_type) { PyErr_Clear(); return 0; }
    }

    if (!datetime_type || !date_type || !timedelta_type)
    {
        PyObject *mod = PyImport_ImportModule("datetime");
        if (!mod) { PyErr_Clear(); return 0; }
        if (!datetime_type)
        {
            datetime_type = PyObject_GetAttrString(mod, "datetime");
            if (!datetime_type) { Py_DECREF(mod); PyErr_Clear(); return 0; }
        }
        if (!date_type)
        {
            date_type = PyObject_GetAttrString(mod, "date");
            if (!date_type) { Py_DECREF(mod); PyErr_Clear(); return 0; }
        }
        if (!timedelta_type)
        {
            timedelta_type = PyObject_GetAttrString(mod, "timedelta");
            if (!timedelta_type) { Py_DECREF(mod); PyErr_Clear(); return 0; }
        }
        Py_DECREF(mod);
    }
    return 1;
}

/* ===== Helper: read variable-length integer (BIFF format) ===== */
static long _read_varint(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t *pos)
{
    if (*pos >= data_len) return -1;
    long value = 0;
    int shift = 0;
    while (1)
    {
        unsigned char b = data[*pos]; (*pos)++;
        value |= (long)(b & 0x7F) << shift;
        if ((b & 0x80) == 0) break;
        shift += 7;
        if (shift > 28 || *pos >= data_len) return -1;
    }
    return value;
}

static long _read_uint32(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos)
{
    if (pos + 4 > data_len) return -1;
    return (long)data[pos] | ((long)data[pos+1] << 8) |
           ((long)data[pos+2] << 16) | ((long)data[pos+3] << 24);
}

static int _read_uint32_safe(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos, long *out)
{
    if (pos + 4 > data_len) return 0;
    *out = (long)data[pos] | ((long)data[pos+1] << 8) |
           ((long)data[pos+2] << 16) | ((long)data[pos+3] << 24);
    return 1;
}

static double _read_double(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos)
{
    if (pos + 8 > data_len) return 0.0;
    unsigned long long bits = 0;
    for (int i = 0; i < 8; i++)
        bits |= (unsigned long long)data[pos + i] << (i * 8);
    double result;
    memcpy(&result, &bits, 8);
    return result;
}

static double _read_rk(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos)
{
    if (pos + 4 > data_len) return 0.0;
    unsigned int raw = (unsigned int)data[pos] | ((unsigned int)data[pos+1] << 8) |
                       ((unsigned int)data[pos+2] << 16) | ((unsigned int)data[pos+3] << 24);
    double result;
    if (raw & 2)
    {
        int int_val = (int)(raw >> 2);
        if (int_val & 0x20000000)
            int_val |= (int)0xC0000000;
        result = (double)int_val;
    }
    else
    {
        unsigned long long mant = (unsigned long long)(raw & 0xFFFFFFFC) << 32;
        double dval;
        memcpy(&dval, &mant, 8);
        result = dval;
    }
    if (raw & 1)
        result /= 100.0;
    return result;
}

static PyObject *_utf16le_to_pystr(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos, long char_len)
{
    if (char_len <= 0) return PyUnicode_FromStringAndSize("", 0);
    Py_ssize_t byte_len = char_len * 2;
    if (pos + byte_len > data_len) return PyUnicode_FromStringAndSize("", 0);
    return PyUnicode_DecodeUTF16((const char *)(data + pos), (Py_ssize_t)byte_len, "little", NULL);
}

/* ===== Read a single cell's value from a worksheet cell record ===== */
static PyObject *_parse_cell_value(const unsigned char *data, Py_ssize_t data_len,
                                    Py_ssize_t body_start, long rec_id, long rec_len,
                                    PyObject *shared_strings, Py_ssize_t ss_len,
                                    PyObject *styles_list, Py_ssize_t styles_len,
                                    PyObject *date_num_fmts, long xf_idx, int *has_cell)
{
    PyObject *cell_val = NULL;
    int is_date = 0;

    if (rec_id == 0x01 || rec_id == 0x03 || rec_id == 0x0B)
    {
        cell_val = Py_None;
    }
    else if (rec_id == 0x02)
    {
        double dval = _read_rk(data, data_len, body_start + 8);
        if (xf_idx > 0 && xf_idx < styles_len)
        {
            PyObject *style = PyList_GetItem(styles_list, xf_idx);
            if (style && PyDict_Check(style))
            {
                PyObject *nf = PyDict_GetItemString(style, "numFmtId");
                if (nf)
                    is_date = (PySet_Contains(date_num_fmts, nf) == 1);
            }
        }
        if (is_date)
        {
            _ensure_types();
            PyObject *epoch = PyDateTime_FromDateAndTime(1899, 12, 30, 0, 0, 0, 0);
            if (!epoch) { PyErr_Clear(); is_date = 0; }
            else
            {
                PyObject *td_kw = Py_BuildValue("{s:d}", "days", dval);
                if (!td_kw) { PyErr_Clear(); Py_DECREF(epoch); is_date = 0; }
                else
                {
                    PyObject *noargs = PyTuple_New(0);
                    PyObject *td = PyObject_Call(timedelta_type, noargs, td_kw);
                    Py_DECREF(noargs);
                    Py_DECREF(td_kw);
                    if (!td) { PyErr_Clear(); Py_DECREF(epoch); is_date = 0; }
                    else
                    {
                        cell_val = PyNumber_Add(epoch, td);
                        Py_DECREF(td);
                        Py_DECREF(epoch);
                        if (!cell_val) { PyErr_Clear(); is_date = 0; }
                    }
                }
            }
        }
        if (!is_date)
        {
            long long_val = (long)dval;
            if (fabs((double)long_val - dval) < 1e-10)
                cell_val = PyLong_FromLong(long_val);
            else
                cell_val = PyFloat_FromDouble(dval);
        }
    }
    else if (rec_id == 0x04 || rec_id == 0x0A)
    {
        if (rec_len >= 9)
            cell_val = (data[body_start + 8] != 0) ? Py_True : Py_False;
        else
            cell_val = Py_False;
    }
    else if (rec_id == 0x05 || rec_id == 0x09)
    {
        double dval = _read_double(data, data_len, body_start + 8);
        if (xf_idx > 0 && xf_idx < styles_len)
        {
            PyObject *style = PyList_GetItem(styles_list, xf_idx);
            if (style && PyDict_Check(style))
            {
                PyObject *nf = PyDict_GetItemString(style, "numFmtId");
                if (nf)
                    is_date = (PySet_Contains(date_num_fmts, nf) == 1);
            }
        }
        if (is_date)
        {
            _ensure_types();
            PyObject *epoch = PyDateTime_FromDateAndTime(1899, 12, 30, 0, 0, 0, 0);
            if (!epoch) { PyErr_Clear(); is_date = 0; }
            else
            {
                PyObject *td_kw = Py_BuildValue("{s:d}", "days", dval);
                if (!td_kw) { PyErr_Clear(); Py_DECREF(epoch); is_date = 0; }
                else
                {
                    PyObject *noargs = PyTuple_New(0);
                    PyObject *td = PyObject_Call(timedelta_type, noargs, td_kw);
                    Py_DECREF(noargs);
                    Py_DECREF(td_kw);
                    if (!td) { PyErr_Clear(); Py_DECREF(epoch); is_date = 0; }
                    else
                    {
                        cell_val = PyNumber_Add(epoch, td);
                        Py_DECREF(td);
                        Py_DECREF(epoch);
                        if (!cell_val) { PyErr_Clear(); is_date = 0; }
                    }
                }
            }
        }
        if (!is_date)
        {
            long long_val = (long)dval;
            if (fabs((double)long_val - dval) < 1e-10)
                cell_val = PyLong_FromLong(long_val);
            else
                cell_val = PyFloat_FromDouble(dval);
        }
    }
    else if (rec_id == 0x06 || rec_id == 0x08)
    {
        if (rec_len >= 12)
        {
            long str_len = _read_uint32(data, data_len, body_start + 8);
            cell_val = _utf16le_to_pystr(data, data_len, body_start + 12, str_len);
        }
        else
        {
            cell_val = PyUnicode_FromStringAndSize("", 0);
        }
    }
    else if (rec_id == 0x07)
    {
        if (rec_len >= 12)
        {
            long idx = _read_uint32(data, data_len, body_start + 8);
            if (idx >= 0 && idx < ss_len)
            {
                cell_val = PyList_GetItem(shared_strings, (Py_ssize_t)idx);
                if (cell_val) Py_INCREF(cell_val);
            }
        }
        if (!cell_val) cell_val = Py_None;
    }

    if (!cell_val) cell_val = Py_None;
    if (cell_val == Py_None || cell_val == Py_True || cell_val == Py_False)
        *has_cell = 0;
    else
        *has_cell = 1;
    return cell_val;
}

/* ===== Clean up cur_cells when _build_row fails ===== */
static void _free_cur_cells(PyObject **cur_cells, long cur_cells_len)
{
    if (!cur_cells) return;
    for (long c = 0; c < cur_cells_len; c++)
    {
        PyObject *v = cur_cells[c];
        if (v && v != Py_None && v != Py_True && v != Py_False)
            Py_DECREF(v);
    }
    free(cur_cells);
}

/* ===== Build a row list from cur_cells array ===== */
static PyObject *_build_row(PyObject **cur_cells, long cur_cells_len, long cols_per_row)
{
    long out_cols = (cols_per_row > 0) ? cols_per_row : (cur_cells_len > 0 ? cur_cells_len : 1);
    PyObject *row_list = PyList_New(out_cols);
    if (!row_list) return NULL;
    for (long c = 0; c < out_cols; c++)
    {
        if (c < cur_cells_len && cur_cells[c])
            PyList_SET_ITEM(row_list, c, cur_cells[c]);
        else
            PyList_SET_ITEM(row_list, c, Py_None);
    }
    return row_list;
}

/* ===== Existing writer functions ===== */

static double _get_excel_serial(PyObject *cell, int cell_kind)
{
    PyObject *ymd = PyObject_GetAttrString(cell, "year");
    if (!ymd) { PyErr_Clear(); return -1.0; }
    long year = PyLong_AsLong(ymd);
    Py_DECREF(ymd);
    if (year < 1900 || year > 9999) return -1.0;

    PyObject *ord = PyObject_CallMethod(cell, "toordinal", NULL);
    if (!ord) { PyErr_Clear(); return -1.0; }
    long ord_val = PyLong_AsLong(ord);
    Py_DECREF(ord);

    long epoch_ord = 693594;
    double serial = (double)(ord_val - epoch_ord);

    if (cell_kind == 1)
    {
        struct { const char *name; double val; } tp[4] = {
            {"hour", 0.0}, {"minute", 0.0}, {"second", 0.0}, {"microsecond", 0.0}
        };
        int ok = 1;
        for (int i = 0; i < 4; i++)
        {
            PyObject *a = PyObject_GetAttrString(cell, tp[i].name);
            if (a)
            {
                tp[i].val = (double)PyLong_AsLong(a);
                if (tp[i].val == -1 && PyErr_Occurred()) { PyErr_Clear(); ok = 0; }
                Py_DECREF(a);
            }
            else { PyErr_Clear(); ok = 0; }
            if (!ok) break;
        }
        if (ok)
        {
            double ts = tp[0].val * 3600.0 + tp[1].val * 60.0 + tp[2].val + tp[3].val / 1e6;
            serial += ts / 86400.0;
        }
    }
    return serial;
}

static PyObject *C_calc_column_widths(PyObject *self, PyObject *args)
{
    PyObject *rows;
    int max_cols;

    if (!PyArg_ParseTuple(args, "Oi", &rows, &max_cols))
        return NULL;
    if (!PyList_Check(rows))
    {
        PyErr_SetString(PyExc_TypeError, "rows must be a list");
        return NULL;
    }

    _ensure_types();

    Py_ssize_t nrows = PyList_Size(rows);
    PyObject *result = PyList_New(max_cols);
    if (!result) return NULL;

    for (int i = 0; i < max_cols; i++)
    {
        double max_width = 0.0;
        for (Py_ssize_t r = 0; r < nrows; r++)
        {
            PyObject *row = PyList_GetItem(rows, r);
            if (!row || !PyList_Check(row)) continue;
            Py_ssize_t row_len = PyList_Size(row);
            if (i >= row_len) continue;

            PyObject *cell = PyList_GetItem(row, i);
            if (!cell || cell == Py_None) continue;

            double temp_width;
            int is_dt = datetime_type ? PyObject_IsInstance(cell, datetime_type) : 0;
            if (is_dt == 1)
                temp_width = 18.0;
            else
            {
                PyObject *str_cell = PyObject_Str(cell);
                if (!str_cell) { PyErr_Clear(); continue; }
                const char *str_data = PyUnicode_AsUTF8(str_cell);
                if (!str_data) { Py_DECREF(str_cell); PyErr_Clear(); continue; }
                const char *nl = strchr(str_data, '\n');
                Py_ssize_t lenn = nl ? (nl - str_data) : (Py_ssize_t)strlen(str_data);
                temp_width = 1.25 * (double)lenn + 4.0;
                Py_DECREF(str_cell);
            }
            if (temp_width > max_width) max_width = temp_width;
        }
        if (max_width == 0.0) max_width = 4.0;
        if (max_width > 255.0) max_width = 255.0;
        PyObject *wv = PyFloat_FromDouble(max_width);
        if (!wv) { Py_DECREF(result); return NULL; }
        PyList_SET_ITEM(result, i, wv);
    }
    return result;
}

/* Shared strings in encoder */
static int _write_str_cell(unsigned char **buf, size_t *buf_size, size_t *buf_cap,
                           PyObject *cell_str, PyObject *ss_dict, PyObject *ss_list,
                           unsigned int col_idx, unsigned int row_idx,
                           unsigned int *sst_uniq, unsigned int *sst_all)
{
    PyObject *idx_obj = PyDict_GetItem(ss_dict, cell_str);
    unsigned int string_index;
    unsigned int style_ref = (row_idx == 0) ? 3 : 0;

    if (idx_obj)
        string_index = (unsigned int)PyLong_AsLong(idx_obj);
    else
    {
        string_index = *sst_uniq;
        PyObject *si = PyLong_FromLong((long)string_index);
        if (!si) return -1;
        if (PyDict_SetItem(ss_dict, cell_str, si) < 0) { Py_DECREF(si); return -1; }
        Py_DECREF(si);
        if (PyList_Append(ss_list, cell_str) < 0) return -1;
        (*sst_uniq)++;
    }
    (*sst_all)++;

    size_t need = *buf_size + 14;
    if (need > *buf_cap)
    {
        *buf_cap = need + 4096;
        unsigned char *nb = (unsigned char *)PyMem_Realloc(*buf, *buf_cap);
        if (!nb) return -1;
        *buf = nb;
    }

    (*buf)[(*buf_size)++] = 0x07;
    (*buf)[(*buf_size)++] = 0x0C;
    memcpy(*buf + *buf_size, &col_idx, 4); (*buf_size) += 4;
    memcpy(*buf + *buf_size, &style_ref, 4); (*buf_size) += 4;
    memcpy(*buf + *buf_size, &string_index, 4); (*buf_size) += 4;
    return 0;
}

static PyObject *C_encode_xlsb_row(PyObject *self, PyObject *args)
{
    PyObject *row;
    PyObject *ss_dict;
    PyObject *ss_list;
    unsigned int sst_uniq;
    unsigned int sst_all;
    unsigned int row_idx;

    if (!PyArg_ParseTuple(args, "OOOIII", &row, &ss_dict, &ss_list, &sst_uniq, &sst_all, &row_idx))
        return NULL;

    if (!PyList_Check(row)) { PyErr_SetString(PyExc_TypeError, "row must be a list"); return NULL; }
    if (!PyDict_Check(ss_dict)) { PyErr_SetString(PyExc_TypeError, "ss_dict must be a dict"); return NULL; }
    if (!PyList_Check(ss_list)) { PyErr_SetString(PyExc_TypeError, "ss_list must be a list"); return NULL; }

    _ensure_types();

    unsigned char *buf = NULL;
    size_t buf_size = 0;
    size_t buf_cap = 4096;
    int err = 0;

    buf = (unsigned char *)PyMem_Malloc(buf_cap);
    if (!buf) return PyErr_NoMemory();

    Py_ssize_t row_len = PyList_Size(row);
    unsigned int last_col = (row_len > 0) ? (unsigned int)(row_len - 1) : 0;

    buf[buf_size++] = 0x00; buf[buf_size++] = 0x19;
    memcpy(buf + buf_size, &row_idx, 4); buf_size += 4;
    memset(buf + buf_size, 0, 4); buf_size += 4;
    unsigned int row_h = 0x12c;
    memcpy(buf + buf_size, &row_h, 4); buf_size += 4;
    buf[buf_size++] = 0x00; buf[buf_size++] = 0x01; buf[buf_size++] = 0x00; buf[buf_size++] = 0x00; buf[buf_size++] = 0x00;
    unsigned int r0 = 0;
    memcpy(buf + buf_size, &r0, 4); buf_size += 4;
    memcpy(buf + buf_size, &last_col, 4); buf_size += 4;

    for (Py_ssize_t col_idx = 0; col_idx < row_len; col_idx++)
    {
        PyObject *cell = PyList_GetItem(row, col_idx);
        if (!cell || cell == Py_None) continue;

        size_t need = buf_size + 16;
        if (need > buf_cap)
        {
            buf_cap = need + 4096;
            unsigned char *nb = (unsigned char *)PyMem_Realloc(buf, buf_cap);
            if (!nb) { err = 1; break; }
            buf = nb;
        }

        unsigned int ci = (unsigned int)col_idx;
        unsigned int zero = 0;

        if (PyUnicode_CheckExact(cell))
        {
            if (_write_str_cell(&buf, &buf_size, &buf_cap, cell, ss_dict, ss_list, ci, row_idx, &sst_uniq, &sst_all) < 0) { err = 1; break; }
        }
        else if (cell == Py_True || cell == Py_False)
        {
            buf[buf_size++] = 0x04; buf[buf_size++] = 0x09;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            memcpy(buf + buf_size, &zero, 4); buf_size += 4;
            buf[buf_size++] = (cell == Py_True) ? 1 : 0;
        }
        else if (PyLong_CheckExact(cell))
        {
            long long val = PyLong_AsLongLong(cell);
            if (val == -1 && PyErr_Occurred()) { PyErr_Clear(); goto fallback; }
            if (val >= -536870912LL && val <= 536870911LL)
            {
                int rk_val = ((int)val << 2) | 2;
                buf[buf_size++] = 0x02; buf[buf_size++] = 0x0C;
                memcpy(buf + buf_size, &ci, 4); buf_size += 4;
                memcpy(buf + buf_size, &zero, 4); buf_size += 4;
                memcpy(buf + buf_size, &rk_val, 4); buf_size += 4;
            }
            else
            {
                buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
                memcpy(buf + buf_size, &ci, 4); buf_size += 4;
                memcpy(buf + buf_size, &zero, 4); buf_size += 4;
                double fval = (double)val;
                memcpy(buf + buf_size, &fval, 8); buf_size += 8;
            }
        }
        else if (PyFloat_CheckExact(cell))
        {
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            memcpy(buf + buf_size, &zero, 4); buf_size += 4;
            double fval = PyFloat_AS_DOUBLE(cell);
            memcpy(buf + buf_size, &fval, 8); buf_size += 8;
        }
        else if (decimal_type && PyObject_IsInstance(cell, decimal_type) == 1)
        {
            PyObject *f = PyObject_CallMethod(cell, "__float__", NULL);
            if (!f) { PyErr_Clear(); goto fallback; }
            double fval = PyFloat_AsDouble(f); Py_DECREF(f);
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            memcpy(buf + buf_size, &zero, 4); buf_size += 4;
            memcpy(buf + buf_size, &fval, 8); buf_size += 8;
        }
        else if (datetime_type && PyObject_IsInstance(cell, datetime_type) == 1)
        {
            double serial = _get_excel_serial(cell, 1);
            if (serial == -1.0) goto fallback;
            unsigned int sv = 1;
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            memcpy(buf + buf_size, &sv, 4); buf_size += 4;
            memcpy(buf + buf_size, &serial, 8); buf_size += 8;
        }
        else if (date_type && PyObject_IsInstance(cell, date_type) == 1)
        {
            double serial = _get_excel_serial(cell, 2);
            if (serial == -1.0) goto fallback;
            unsigned int sv = 2;
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            memcpy(buf + buf_size, &sv, 4); buf_size += 4;
            memcpy(buf + buf_size, &serial, 8); buf_size += 8;
        }
        else
        {
            fallback:
            {
                PyObject *str_cell = PyObject_Str(cell);
                if (!str_cell) { PyErr_Clear(); continue; }
                if (_write_str_cell(&buf, &buf_size, &buf_cap, str_cell, ss_dict, ss_list, ci, row_idx, &sst_uniq, &sst_all) < 0)
                { Py_DECREF(str_cell); err = 1; break; }
                Py_DECREF(str_cell);
            }
        }
    }

    if (err) { PyMem_Free(buf); return PyErr_NoMemory(); }

    PyObject *result_bytes = PyBytes_FromStringAndSize((const char *)buf, buf_size);
    PyMem_Free(buf);
    if (!result_bytes) return NULL;

    return Py_BuildValue("OII", result_bytes, sst_uniq, sst_all);
}

/* ===== Streaming XLSB worksheet reader (iterator type) ===== */

/* Forward declaration for iternext */
static PyObject *XlsbReader_iternext(PyObject *self);

typedef struct {
    PyObject_HEAD
    /* Data buffer */
    Py_buffer view;
    const unsigned char *data;
    Py_ssize_t data_len;
    Py_ssize_t pos;
    int done;

    /* Lookup structures */
    PyObject *shared_strings;
    Py_ssize_t ss_len;
    PyObject *styles_list;
    Py_ssize_t styles_len;
    PyObject *date_num_fmts;

    /* Per-row cell buffer */
    PyObject **cur_cells;
    long cur_cells_len;

    /* State tracking */
    long cur_row_num;          /* row number of currently accumulating row */
    long gap_rows_remaining;   /* number of empty rows to emit before next data row */
    int is_first_row;          /* 1 during first row processing (column alignment) */
    long first_col;
    long cols_per_row;
    int is_initialized;        /* 1 after first Row record consumed */
} XlsbReader;

static void XlsbReader_dealloc(XlsbReader *self)
{
    if (self->cur_cells)
    {
        /* Cells have been stolen by PyList_SET_ITEM when rows were emitted,
           OR haven't been emitted yet. We only need to free the array itself;
           any remaining cells are owned by the row list or are stale pointers. */
        free(self->cur_cells);
    }
    if (self->view.buf)
        PyBuffer_Release(&self->view);
    Py_XDECREF(self->shared_strings);
    Py_XDECREF(self->styles_list);
    Py_XDECREF(self->date_num_fmts);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyTypeObject XlsbReader_Type = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name = "_c_core.XlsbReader",
    .tp_basicsize = sizeof(XlsbReader),
    .tp_dealloc = (destructor)XlsbReader_dealloc,
    .tp_flags = Py_TPFLAGS_DEFAULT,
    .tp_doc = "Streaming XLSB worksheet reader (iterator)",
    .tp_iter = PyObject_SelfIter,
    .tp_iternext = XlsbReader_iternext,
};

static PyObject *_build_row_and_free_cells(XlsbReader *self)
{
    long out_cols = (self->cols_per_row > 0) ? self->cols_per_row :
                    (self->cur_cells_len > 0 ? self->cur_cells_len : 1);
    PyObject *row_list = PyList_New(out_cols);
    if (!row_list)
    {
        _free_cur_cells(self->cur_cells, self->cur_cells_len);
        self->cur_cells = NULL; self->cur_cells_len = 0;
        return NULL;
    }
    for (long c = 0; c < out_cols; c++)
    {
        if (c < self->cur_cells_len && self->cur_cells[c])
            PyList_SET_ITEM(row_list, c, self->cur_cells[c]);
        else
            PyList_SET_ITEM(row_list, c, Py_None);
    }
    /* Cells stolen by PyList_SET_ITEM - just free the array */
    free(self->cur_cells);
    self->cur_cells = NULL;
    self->cur_cells_len = 0;
    return row_list;
}

static PyObject *_make_empty_row(XlsbReader *self)
{
    long cols = (self->cols_per_row > 0) ? self->cols_per_row : 1;
    PyObject *row = PyList_New(cols);
    if (!row) return NULL;
    for (long c = 0; c < cols; c++)
        PyList_SET_ITEM(row, c, Py_None);
    return row;
}

/* Helper: read a cell record and store in cur_cells */
static int _process_cell_record(XlsbReader *self, const unsigned char *data,
                                 Py_ssize_t data_len, Py_ssize_t body_start,
                                 long rec_id, long rec_len)
{
    long col, xf_idx;
    if (!_read_uint32_safe(data, data_len, body_start, &col)) return 1;
    if (!_read_uint32_safe(data, data_len, body_start + 4, &xf_idx)) return 1;
    xf_idx &= 0xFFFFFF;

    /* On first row, track column range for alignment (frozen after first row) */
    if (self->is_first_row)
    {
        if (self->first_col < 0 || col < self->first_col)
            self->first_col = col;
        long observed_max = col;
        long new_cols = observed_max - self->first_col + 1;
        if (new_cols > self->cols_per_row)
            self->cols_per_row = new_cols;
    }

    int has_cell_alloc;
    PyObject *cell_val = _parse_cell_value(data, data_len, body_start,
                                             rec_id, rec_len,
                                             self->shared_strings, self->ss_len,
                                             self->styles_list, self->styles_len,
                                             self->date_num_fmts, xf_idx, &has_cell_alloc);
    if (!cell_val) return 1;

    /* Index relative to frozen first_col (if first_col < 0, default to col itself) */
    long fc = self->first_col;
    if (fc < 0) fc = col;
    long idx = col - fc;
    if (idx < 0) idx = 0;

    if (idx >= self->cur_cells_len)
    {
        PyObject **new_cells = realloc(self->cur_cells, sizeof(PyObject*) * (idx + 1));
        if (!new_cells)
        {
            if (has_cell_alloc) Py_DECREF(cell_val);
            return 0;
        }
        for (long c = self->cur_cells_len; c <= idx; c++)
            new_cells[c] = NULL;
        self->cur_cells = new_cells;
        self->cur_cells_len = idx + 1;
    }
    Py_XDECREF(self->cur_cells[idx]);
    self->cur_cells[idx] = cell_val;
    return 1;
}

static PyObject *XlsbReader_iternext(PyObject *self_ptr)
{
    XlsbReader *self = (XlsbReader *)self_ptr;
    if (self->done) return NULL;

    const unsigned char *data = self->data;
    Py_ssize_t data_len = self->data_len;

    /* Phase 0: Emit pending gap rows (one per __next__ call) */
    if (self->gap_rows_remaining > 0)
    {
        self->gap_rows_remaining--;
        PyObject *row = _make_empty_row(self);
        if (!row) { self->done = 1; return NULL; }
        return row;
    }

    /* Phase 1: On first call, skip past metadata to first Row record */
    if (!self->is_initialized)
    {
        while (self->pos < data_len)
        {
            long rec_id = _read_varint(data, data_len, &self->pos);
            if (rec_id < 0) { self->done = 1; return NULL; }
            long rec_len = _read_varint(data, data_len, &self->pos);
            if (rec_len < 0) { self->done = 1; return NULL; }
            Py_ssize_t body_start = self->pos;
            self->pos += rec_len;
            if (rec_id == 0x00)
            {
                long new_row;
                _read_uint32_safe(data, data_len, body_start, &new_row);
                self->cur_row_num = new_row;
                self->is_initialized = 1;
                break;
            }
        }
        if (!self->is_initialized) { self->done = 1; return NULL; }
    }

    /* Phase 2: Read cell records until next Row record or EOF */
    while (self->pos < data_len)
    {
        long rec_id = _read_varint(data, data_len, &self->pos);
        if (rec_id < 0) break;
        long rec_len = _read_varint(data, data_len, &self->pos);
        if (rec_len < 0) break;

        Py_ssize_t body_start = self->pos;
        self->pos += rec_len;

        if (rec_id == 0x00)
        {
            /* Next Row record: build and return the accumulated row */
            PyObject *row = _build_row_and_free_cells(self);
            if (!row) { self->done = 1; return NULL; }

            long new_row;
            _read_uint32_safe(data, data_len, body_start, &new_row);

            /* Calculate gap rows between previous and new row */
            self->gap_rows_remaining = new_row - self->cur_row_num - 1;

            if (self->is_first_row) self->is_first_row = 0;
            self->cur_row_num = new_row;
            return row;
        }
        else if (rec_id >= 0x01 && rec_id <= 0x0B)
        {
            if (!_process_cell_record(self, data, data_len, body_start, rec_id, rec_len))
            { self->done = 1; return NULL; }
        }
        else if (rec_id == 0x92)
        {
            break;
        }
    }

    /* Phase 3: EOF - emit last row */
    self->done = 1;
    if (self->cur_cells && self->cur_cells_len > 0)
        return _build_row_and_free_cells(self);
    if (self->cur_row_num >= 0)
        return _make_empty_row(self);
    return NULL;
}

/* ===== Factory function ===== */
static PyObject *C_read_xlsb_worksheet(PyObject *self_mod, PyObject *args)
{
    Py_buffer view;
    PyObject *shared_strings;
    PyObject *styles_list;
    PyObject *date_num_fmts;

    if (!PyArg_ParseTuple(args, "y*OOO", &view, &shared_strings, &styles_list, &date_num_fmts))
        return NULL;

    if (!PyList_Check(shared_strings)) { PyErr_SetString(PyExc_TypeError, "shared_strings must be a list"); PyBuffer_Release(&view); return NULL; }
    if (!PyList_Check(styles_list)) { PyErr_SetString(PyExc_TypeError, "styles_list must be a list"); PyBuffer_Release(&view); return NULL; }
    if (!PySet_Check(date_num_fmts)) { PyErr_SetString(PyExc_TypeError, "date_num_fmts must be a set"); PyBuffer_Release(&view); return NULL; }

    _ensure_types();

    XlsbReader *reader = PyObject_New(XlsbReader, &XlsbReader_Type);
    if (!reader) { PyBuffer_Release(&view); return NULL; }

    memcpy(&reader->view, &view, sizeof(Py_buffer));
    reader->data = (const unsigned char *)view.buf;
    reader->data_len = (Py_ssize_t)view.len;
    reader->pos = 0;
    reader->done = 0;

    reader->shared_strings = shared_strings;
    reader->ss_len = PyList_Size(shared_strings);
    reader->styles_list = styles_list;
    reader->styles_len = PyList_Size(styles_list);
    reader->date_num_fmts = date_num_fmts;

    reader->cur_cells = NULL;
    reader->cur_cells_len = 0;
    reader->cur_row_num = -1;
    reader->gap_rows_remaining = 0;
    reader->is_first_row = 1;
    reader->first_col = -1;
    reader->cols_per_row = -1;
    reader->is_initialized = 0;

    Py_INCREF(shared_strings);
    Py_INCREF(styles_list);
    Py_INCREF(date_num_fmts);

    return (PyObject *)reader;
}

/* ===== Module definition ===== */
static PyMethodDef CMethods[] = {
    {"read_xlsb_worksheet", C_read_xlsb_worksheet, METH_VARARGS,
     "Read an XLSB worksheet from raw bytes. Returns an iterator of rows.\n\n"
     "Args: (data, shared_strings, styles_list, date_num_fmts)\n"
     "Yields: rows one at a time"},
    {"encode_xlsb_row", C_encode_xlsb_row, METH_VARARGS,
     "Encode a single XLSB row. Returns (bytes, sst_uniq, sst_all)."},
    {"calc_column_widths", C_calc_column_widths, METH_VARARGS,
     "Calculate column widths from row data."},
    {NULL, NULL, 0, NULL}
};

static struct PyModuleDef c_core_module = {
    PyModuleDef_HEAD_INIT,
    "_c_core",
    "C-optimized functions for xlspy",
    -1,
    CMethods
};

PyMODINIT_FUNC PyInit__c_core(void)
{
    PyDateTime_IMPORT;
    if (PyType_Ready(&XlsbReader_Type) < 0)
        return NULL;
    PyObject *m = PyModule_Create(&c_core_module);
    return m;
}
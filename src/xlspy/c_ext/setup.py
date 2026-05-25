from setuptools import Extension, setup

module = Extension(
    '_c_core',
    sources=['_c_core.c'],
)

setup(
    name='xlspy_c_ext',
    ext_modules=[module],
)


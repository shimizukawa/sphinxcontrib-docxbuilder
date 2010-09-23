Sphinx docx builder extension.

Features
========

* Generate docx file as Sphinx builder.

Setup
=====

by buildout
------------
Make environment::

    $ python bootstrap.py -d init
    $ bin/buildout

Usage
=====

Execute sphinx-build with below option::

    $ bin/sphinx-build -b docx [input-dir] [output-dir]
    $ ls [output-dir]
    output.docx

for example building, simply run below::

    $ bin/example
    ...
    Saved new file to: examples/index.docx


Requirements
============

* Python 2.6 or later (not support 3.x)
* python-docx (not released, but included)
* setuptools or distriubte.

History
=======

0.0.1 (unreleased)
--------------------
* first pre-alpha release



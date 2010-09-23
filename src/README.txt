Sphinx docx builder extension.

Features
========

* Generate docx file as Sphinx builder.

Setup
=====

by easy_install
----------------
Make environment::

   $ easy_install sphinxcontrib-docxbuilder

Usage
=====

Execute sphinx-build with below option::

   $ sphinx-build -b docx [input-dir] [output-dir]
   $ ls [output-dir]
   output.docx


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



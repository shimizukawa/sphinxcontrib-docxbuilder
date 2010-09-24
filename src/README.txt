Sphinx docx builder extension generate single docx file from Sphinx document
source. This extension use python-docx module (included) for the docx file
generation.

Features
========

* This extension work on Multi-platform (not need OpenOffice or MS Word).
* Usable sphinx syntax and directives:
    * heading line output
    * paragraph output (standard body text)
    * image and figure directive output
    * bullet-list and numbered-list output
    * table output (restrictive)

Currently, many directives and indented block are not work correctly, yet.

Setup
=====

Make environment by easy_install
---------------------------------

Not yet.

Make environment by buildout
-----------------------------

'hg clone' or download sphinxcontrib-docxbuilder archive from 'get source'
menu at http://bitbucket.org/shimizukawa/sphinxcontrib-docxbuilder ::

    $ cd /path/to/sphinxcontrib-docxbuilder
    $ python bootstrap.py -d init
    $ bin/buildout


run example
------------

for example sphinx-docx building, simply run below::

    $ bin/example
    ...
    Saved new file to: examples/example-0.1.docx


Usage
=====

Set 'sphinxcontrib-docxbuilder' to 'extensions' line of target sphinx source
conf.py::

    extensions = ['sphinxcontrib-docxbuilder']

Execute sphinx-build with below option::

    $ bin/sphinx-build -b docx [input-dir] [output-dir]
    $ ls [output-dir]
    output.docx


Requirements
============

* Python 2.6 or later (not support 3.x)
* `python-docx <http://github.com/mikemaccana/python-docx>`_
  (not released, but included), Thanks Mike MacCana.
* setuptools or distriubte.

History
=======

0.0.1 (unreleased)
--------------------
* Not released.



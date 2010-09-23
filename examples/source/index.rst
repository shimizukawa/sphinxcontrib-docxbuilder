=================================
Python docx モジュールへようこそ
=================================

Pythonのコード200行だけでdocxを作り変更するぜ
==============================================
このモジュールが作られたきっかけは、私がPythonでMS Word .doc ファイルを
扱えないかPyPIやStackoverflowで探したことだ。残念ながら見つけることが
出来たのは以下の方法だけだった:

1. COM automation
2. .net or Java
3. Automating OpenOffice or MS Office

For those of us who prefer something simpler, I made docx.

Making documents
=================

The docx module has the following features:

* Paragraphs
* Bullets
* Numbered lists
* Multiple levels of headings
* Tables
* Document Properties

Tables are just lists of lists, like this:

-- -- --
A1 A2 A3
B1 B2 B3
C1 C2 C3
-- -- --

Editing documents
==================

Thanks to the awesomeness of the lxml module, we can:

* Search and replace
* Extract plain text of document
* Add and delete items anywhere within the document

.. figure:: image1.png

    This is a test description


.. .. page-break::

Ideas? Questions? Want to contribute?
======================================

Email <python.docx@librelist.com>


# -*- coding: utf-8 -*-
"""
    sphinxcontrib-docxwriter
    ~~~~~~~~~~~~~~~~~~~~~~~~~~

    Custom docutils writer for OpenXML (docx).

    :copyright: Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""

import re
import textwrap

from docutils import nodes, writers

from sphinx import addnodes
from sphinx.locale import admonitionlabels, versionlabels, _

import docx
import sys


def dprint(**kw):
    f = sys._getframe(1)
    if kw:
        text = ', '.join('%s = %s' % (k,v) for k,v in kw.items())
    else:
        text = dict((k,v) for k,v in f.f_locals.items() if k != 'self')

    try:
        print >>sys.stderr, '###', f.f_code.co_name, text
    except:
        print >>sys.stderr, ''


class DocxWriter(writers.Writer):
    supported = ('docx',)
    settings_spec = ('No options here.', '', ())
    settings_defaults = {}

    output = None

    def __init__(self, builder):
        writers.Writer.__init__(self)
        self.builder = builder
        self.docx_document = docx.newdocument()
        self.docbody = self.docx_document.xpath(
                '/w:document/w:body', namespaces=docx.nsprefixes)[0]

    def save(self, filename):
        relationships = docx.relationshiplist()
        appprops = docx.appproperties()
        contenttypes = docx.contenttypes()
        websettings = docx.websettings()
        wordrelationships = docx.wordrelationships(relationships)
        coreprops = docx.coreproperties(
                title='Python docx demo',
                subject='A practical example of making docx from Python',
                creator='Mike MacCana',
                keywords=['python','Office Open XML','Word'])

        docx.savedocx(self.docx_document, coreprops, appprops, contenttypes,
                websettings, wordrelationships, filename)

    def translate(self):
        visitor = DocxTranslator(self.document, self.builder, self.docbody)
        self.document.walkabout(visitor)
        self.output = '' #visitor.body


class DocxTranslator(nodes.NodeVisitor):
    sectionchars = '*=-~"+`'

    def __init__(self, document, builder, docbody):
        self.builder = builder
        self.docbody = docbody
        nodes.NodeVisitor.__init__(self, document)

        self.states = [[]]
        self.stateindent = [0]
        #self.list_counter = []
        self.sectionlevel = 1
        #self.table = None

    def add_text(self, text):
        dprint()
        self.states[-1].append((-1, text))

    def new_state(self, indent=2):
        dprint()
        self.states.append([])
        self.stateindent.append(indent)

    def end_state(self, wrap=False, end=[''], first=None):
        dprint()
        content = self.states.pop()
        maxindent = sum(self.stateindent)
        indent = self.stateindent.pop()
        result = []
        toformat = []
        def do_format():
            if not toformat:
                return
            #if wrap:
            #    res = textwrap.wrap(''.join(toformat), width=MAXWIDTH-maxindent)
            #else:
            #    res = ''.join(toformat).splitlines()
            res = ''.join(toformat).splitlines()
            if end:
                res += end
            result.append((indent, res))
        for itemindent, item in content:
            if itemindent == -1:
                toformat.append(item)
            else:
                do_format()
                result.append((indent + itemindent, item))
                toformat = []
        do_format()
        if first is not None and result:
            itemindent, item = result[0]
            if item:
                result.insert(0, (itemindent - indent, [first + item[0]]))
                result[1] = (itemindent, item[1:])
        self.states[-1].extend(result)

    def visit_document(self, node):
        dprint()
        self.new_state(0)

    def depart_document(self, node):
        dprint()
        self.end_state()
        #self.body = '\n'.join(line and (' '*indent + line)
        #                      for indent, lines in self.states[0]
        #                      for line in lines)
        ## XXX header/footer?

    def visit_highlightlang(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_section(self, node):
        dprint()
        self.sectionlevel += 1

    def depart_section(self, node):
        dprint()
        self.sectionlevel -= 1

    def visit_topic(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_topic(self, node):
        pass
        dprint()
        #self.end_state()

    visit_sidebar = visit_topic
    depart_sidebar = depart_topic

    def visit_rubric(self, node):
        pass
        dprint()
        #self.new_state(0)
        #self.add_text('-[ ')

    def depart_rubric(self, node):
        pass
        dprint()
        #self.add_text(' ]-')
        #self.end_state()

    def visit_compound(self, node):
        pass
        dprint()
        pass

    def depart_compound(self, node):
        pass
        dprint()
        pass

    def visit_glossary(self, node):
        pass
        dprint()
        pass

    def depart_glossary(self, node):
        pass
        dprint()
        pass

    def visit_title(self, node):
        dprint()
        #if isinstance(node.parent, nodes.Admonition):
        #    self.add_text(node.astext()+': ')
        #    raise nodes.SkipNode
        self.new_state(0)

    def depart_title(self, node):
        dprint()
        text = ''.join(x[1] for x in self.states.pop() if x[0] == -1)
        self.stateindent.pop()
        dprint(text=repr(text), level=self.sectionlevel)
        self.docbody.append(docx.heading(text, self.sectionlevel))

    def visit_subtitle(self, node):
        pass
        dprint()
        pass

    def depart_subtitle(self, node):
        pass
        dprint()
        pass

    def visit_attribution(self, node):
        pass
        dprint()
        #self.add_text('-- ')

    def depart_attribution(self, node):
        pass
        dprint()
        pass

    def visit_desc(self, node):
        pass
        dprint()
        pass

    def depart_desc(self, node):
        pass
        dprint()
        pass

    def visit_desc_signature(self, node):
        pass
        dprint()
        #self.new_state(0)
        #if node.parent['objtype'] in ('class', 'exception'):
        #    self.add_text('%s ' % node.parent['objtype'])

    def depart_desc_signature(self, node):
        pass
        dprint()
        ## XXX: wrap signatures in a way that makes sense
        #self.end_state(wrap=False, end=None)

    def visit_desc_name(self, node):
        pass
        dprint()
        pass

    def depart_desc_name(self, node):
        pass
        dprint()
        pass

    def visit_desc_addname(self, node):
        pass
        dprint()
        pass

    def depart_desc_addname(self, node):
        pass
        dprint()
        pass

    def visit_desc_type(self, node):
        pass
        dprint()
        pass

    def depart_desc_type(self, node):
        pass
        dprint()
        pass

    def visit_desc_returns(self, node):
        pass
        dprint()
        #self.add_text(' -> ')

    def depart_desc_returns(self, node):
        pass
        dprint()
        pass

    def visit_desc_parameterlist(self, node):
        pass
        dprint()
        #self.add_text('(')
        #self.first_param = 1

    def depart_desc_parameterlist(self, node):
        pass
        dprint()
        #self.add_text(')')

    def visit_desc_parameter(self, node):
        pass
        dprint()
        #if not self.first_param:
        #    self.add_text(', ')
        #else:
        #    self.first_param = 0
        #self.add_text(node.astext())
        #raise nodes.SkipNode

    def visit_desc_optional(self, node):
        pass
        dprint()
        #self.add_text('[')

    def depart_desc_optional(self, node):
        pass
        dprint()
        #self.add_text(']')

    def visit_desc_annotation(self, node):
        pass
        dprint()
        pass

    def depart_desc_annotation(self, node):
        pass
        dprint()
        pass

    def visit_refcount(self, node):
        pass
        dprint()
        pass

    def depart_refcount(self, node):
        pass
        dprint()
        pass

    def visit_desc_content(self, node):
        pass
        dprint()
        #self.new_state()
        #self.add_text('\n')

    def depart_desc_content(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_figure(self, node):
        pass
        dprint()
        #self.new_state()

    def depart_figure(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_caption(self, node):
        pass
        dprint()
        pass

    def depart_caption(self, node):
        pass
        dprint()
        pass

    def visit_productionlist(self, node):
        pass
        dprint()
        #self.new_state()
        #names = []
        #for production in node:
        #    names.append(production['tokenname'])
        #maxlen = max(len(name) for name in names)
        #for production in node:
        #    if production['tokenname']:
        #        self.add_text(production['tokenname'].ljust(maxlen) + ' ::=')
        #        lastname = production['tokenname']
        #    else:
        #        self.add_text('%s    ' % (' '*len(lastname)))
        #    self.add_text(production.astext() + '\n')
        #self.end_state(wrap=False)
        #raise nodes.SkipNode

    def visit_seealso(self, node):
        pass
        dprint()
        self.new_state()

    def depart_seealso(self, node):
        pass
        dprint()
        self.end_state(first='')

    def visit_footnote(self, node):
        pass
        dprint()
        #self._footnote = node.children[0].astext().strip()
        #self.new_state(len(self._footnote) + 3)

    def depart_footnote(self, node):
        pass
        dprint()
        #self.end_state(first='[%s] ' % self._footnote)

    def visit_citation(self, node):
        pass
        dprint()
        #if len(node) and isinstance(node[0], nodes.label):
        #    self._citlabel = node[0].astext()
        #else:
        #    self._citlabel = ''
        #self.new_state(len(self._citlabel) + 3)

    def depart_citation(self, node):
        pass
        dprint()
        #self.end_state(first='[%s] ' % self._citlabel)

    def visit_label(self, node):
        pass
        dprint()
        raise nodes.SkipNode

    # XXX: option list could use some better styling

    def visit_option_list(self, node):
        pass
        dprint()
        pass

    def depart_option_list(self, node):
        pass
        dprint()
        pass

    def visit_option_list_item(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_option_list_item(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_option_group(self, node):
        pass
        dprint()
        #self._firstoption = True

    def depart_option_group(self, node):
        pass
        dprint()
        #self.add_text('     ')

    def visit_option(self, node):
        pass
        dprint()
        #if self._firstoption:
        #    self._firstoption = False
        #else:
        #    self.add_text(', ')

    def depart_option(self, node):
        pass
        dprint()
        pass

    def visit_option_string(self, node):
        pass
        dprint()
        pass

    def depart_option_string(self, node):
        pass
        dprint()
        pass

    def visit_option_argument(self, node):
        pass
        dprint()
        #self.add_text(node['delimiter'])

    def depart_option_argument(self, node):
        pass
        dprint()
        pass

    def visit_description(self, node):
        pass
        dprint()
        pass

    def depart_description(self, node):
        pass
        dprint()
        pass

    def visit_tabular_col_spec(self, node):
        pass
        dprint()
        raise nodes.SkipNode

    def visit_colspec(self, node):
        pass
        dprint()
        #self.table[0].append(node['colwidth'])
        raise nodes.SkipNode

    def visit_tgroup(self, node):
        pass
        dprint()
        pass

    def depart_tgroup(self, node):
        pass
        dprint()
        pass

    def visit_thead(self, node):
        pass
        dprint()
        pass

    def depart_thead(self, node):
        pass
        dprint()
        pass

    def visit_tbody(self, node):
        pass
        dprint()
        #self.table.append('sep')

    def depart_tbody(self, node):
        pass
        dprint()
        pass

    def visit_row(self, node):
        pass
        dprint()
        #self.table.append([])

    def depart_row(self, node):
        pass
        dprint()
        pass

    def visit_entry(self, node):
        pass
        dprint()
        #if node.has_key('morerows') or node.has_key('morecols'):
        #    raise NotImplementedError('Column or row spanning cells are '
        #                              'not implemented.')
        #self.new_state(0)

    def depart_entry(self, node):
        pass
        dprint()
        #text = '\n'.join('\n'.join(x[1]) for x in self.states.pop())
        #self.stateindent.pop()
        #self.table[-1].append(text)

    def visit_table(self, node):
        pass
        dprint()
        #if self.table:
        #    raise NotImplementedError('Nested tables are not supported.')
        #self.new_state(0)
        #self.table = [[]]

    def depart_table(self, node):
        pass
        dprint()
        #lines = self.table[1:]
        #fmted_rows = []
        #colwidths = self.table[0]
        #realwidths = colwidths[:]
        #separator = 0
        ## don't allow paragraphs in table cells for now
        #for line in lines:
        #    if line == 'sep':
        #        separator = len(fmted_rows)
        #    else:
        #        cells = []
        #        for i, cell in enumerate(line):
        #            par = textwrap.wrap(cell, width=colwidths[i])
        #            if par:
        #                maxwidth = max(map(len, par))
        #            else:
        #                maxwidth = 0
        #            realwidths[i] = max(realwidths[i], maxwidth)
        #            cells.append(par)
        #        fmted_rows.append(cells)

        #def writesep(char='-'):
        #    out = ['+']
        #    for width in realwidths:
        #        out.append(char * (width+2))
        #        out.append('+')
        #    self.add_text(''.join(out) + '\n')

        #def writerow(row):
        #    lines = map(None, *row)
        #    for line in lines:
        #        out = ['|']
        #        for i, cell in enumerate(line):
        #            if cell:
        #                out.append(' ' + cell.ljust(realwidths[i]+1))
        #            else:
        #                out.append(' ' * (realwidths[i] + 2))
        #            out.append('|')
        #        self.add_text(''.join(out) + '\n')

        #for i, row in enumerate(fmted_rows):
        #    if separator and i == separator:
        #        writesep('=')
        #    else:
        #        writesep('-')
        #    writerow(row)
        #writesep('-')
        #self.table = None
        #self.end_state(wrap=False)

    def visit_acks(self, node):
        pass
        dprint()
        #self.new_state(0)
        #self.add_text(', '.join(n.astext() for n in node.children[0].children)
        #              + '.')
        #self.end_state()
        raise nodes.SkipNode

    def visit_image(self, node):
        pass
        dprint()
        #self.add_text(_('[image]'))
        raise nodes.SkipNode

    def visit_transition(self, node):
        pass
        dprint()
        #indent = sum(self.stateindent)
        #self.new_state(0)
        #self.add_text('=' * (MAXWIDTH - indent))
        #self.end_state()
        raise nodes.SkipNode

    def visit_bullet_list(self, node):
        pass
        dprint()
        #self.list_counter.append(-1)

    def depart_bullet_list(self, node):
        pass
        dprint()
        #self.list_counter.pop()

    def visit_enumerated_list(self, node):
        pass
        dprint()
        #self.list_counter.append(0)

    def depart_enumerated_list(self, node):
        pass
        dprint()
        #self.list_counter.pop()

    def visit_definition_list(self, node):
        pass
        dprint()
        #self.list_counter.append(-2)

    def depart_definition_list(self, node):
        pass
        dprint()
        #self.list_counter.pop()

    def visit_list_item(self, node):
        pass
        dprint()
        #if self.list_counter[-1] == -1:
        #    # bullet list
        #    self.new_state(2)
        #elif self.list_counter[-1] == -2:
        #    # definition list
        #    pass
        #else:
        #    # enumerated list
        #    self.list_counter[-1] += 1
        #    self.new_state(len(str(self.list_counter[-1])) + 2)

    def depart_list_item(self, node):
        pass
        dprint()
        #if self.list_counter[-1] == -1:
        #    self.end_state(first='* ', end=None)
        #elif self.list_counter[-1] == -2:
        #    pass
        #else:
        #    self.end_state(first='%s. ' % self.list_counter[-1], end=None)

    def visit_definition_list_item(self, node):
        pass
        dprint()
        #self._li_has_classifier = len(node) >= 2 and \
        #                          isinstance(node[1], nodes.classifier)

    def depart_definition_list_item(self, node):
        pass
        dprint()
        pass

    def visit_term(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_term(self, node):
        pass
        dprint()
        #if not self._li_has_classifier:
        #    self.end_state(end=None)

    def visit_classifier(self, node):
        pass
        dprint()
        #self.add_text(' : ')

    def depart_classifier(self, node):
        pass
        dprint()
        #self.end_state(end=None)

    def visit_definition(self, node):
        pass
        dprint()
        #self.new_state()

    def depart_definition(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_field_list(self, node):
        pass
        dprint()
        pass

    def depart_field_list(self, node):
        pass
        dprint()
        pass

    def visit_field(self, node):
        pass
        dprint()
        pass

    def depart_field(self, node):
        pass
        dprint()
        pass

    def visit_field_name(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_field_name(self, node):
        pass
        dprint()
        #self.add_text(':')
        #self.end_state(end=None)

    def visit_field_body(self, node):
        pass
        dprint()
        #self.new_state()

    def depart_field_body(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_centered(self, node):
        pass
        dprint()
        pass

    def depart_centered(self, node):
        pass
        dprint()
        pass

    def visit_hlist(self, node):
        pass
        dprint()
        pass

    def depart_hlist(self, node):
        pass
        dprint()
        pass

    def visit_hlistcol(self, node):
        pass
        dprint()
        pass

    def depart_hlistcol(self, node):
        pass
        dprint()
        pass

    def visit_admonition(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_admonition(self, node):
        pass
        dprint()
        #self.end_state()

    def _visit_admonition(self, node):
        pass
        dprint()
        #self.new_state(2)

    def _make_depart_admonition(name):
        pass
        dprint()
        #def depart_admonition(self, node):
        #    self.end_state(first=admonitionlabels[name] + ': ')
        #return depart_admonition

    visit_attention = _visit_admonition
    depart_attention = _make_depart_admonition('attention')
    visit_caution = _visit_admonition
    depart_caution = _make_depart_admonition('caution')
    visit_danger = _visit_admonition
    depart_danger = _make_depart_admonition('danger')
    visit_error = _visit_admonition
    depart_error = _make_depart_admonition('error')
    visit_hint = _visit_admonition
    depart_hint = _make_depart_admonition('hint')
    visit_important = _visit_admonition
    depart_important = _make_depart_admonition('important')
    visit_note = _visit_admonition
    depart_note = _make_depart_admonition('note')
    visit_tip = _visit_admonition
    depart_tip = _make_depart_admonition('tip')
    visit_warning = _visit_admonition
    depart_warning = _make_depart_admonition('warning')

    def visit_versionmodified(self, node):
        pass
        dprint()
        #self.new_state(0)
        #if node.children:
        #    self.add_text(versionlabels[node['type']] % node['version'] + ': ')
        #else:
        #    self.add_text(versionlabels[node['type']] % node['version'] + '.')

    def depart_versionmodified(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_literal_block(self, node):
        pass
        dprint()
        #self.new_state()

    def depart_literal_block(self, node):
        pass
        dprint()
        #self.end_state(wrap=False)

    def visit_doctest_block(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_doctest_block(self, node):
        pass
        dprint()
        #self.end_state(wrap=False)

    def visit_line_block(self, node):
        pass
        dprint()
        #self.new_state(0)

    def depart_line_block(self, node):
        pass
        dprint()
        #self.end_state(wrap=False)

    def visit_line(self, node):
        pass
        dprint()
        pass

    def depart_line(self, node):
        pass
        dprint()
        pass

    def visit_block_quote(self, node):
        pass
        dprint()
        #self.new_state()

    def depart_block_quote(self, node):
        pass
        dprint()
        #self.end_state()

    def visit_compact_paragraph(self, node):
        pass
        dprint()
        pass

    def depart_compact_paragraph(self, node):
        pass
        dprint()
        pass

    def visit_paragraph(self, node):
        dprint()
        #if not isinstance(node.parent, nodes.Admonition) or \
        #       isinstance(node.parent, addnodes.seealso):
        #    self.new_state(0)
        self.docbody.append(docx.paragraph(node.astext()))

    def depart_paragraph(self, node):
        dprint()
        #if not isinstance(node.parent, nodes.Admonition) or \
        #       isinstance(node.parent, addnodes.seealso):
        #    self.end_state()

    def visit_target(self, node):
        pass
        dprint()
        raise nodes.SkipNode

    def visit_index(self, node):
        pass
        dprint()
        raise nodes.SkipNode

    def visit_substitution_definition(self, node):
        pass
        dprint()
        raise nodes.SkipNode

    def visit_pending_xref(self, node):
        pass
        dprint()
        pass

    def depart_pending_xref(self, node):
        pass
        dprint()
        pass

    def visit_reference(self, node):
        pass
        dprint()
        pass

    def depart_reference(self, node):
        pass
        dprint()
        pass

    def visit_download_reference(self, node):
        pass
        dprint()
        pass

    def depart_download_reference(self, node):
        pass
        dprint()
        pass

    def visit_emphasis(self, node):
        pass
        dprint()
        #self.add_text('*')

    def depart_emphasis(self, node):
        pass
        dprint()
        #self.add_text('*')

    def visit_literal_emphasis(self, node):
        pass
        dprint()
        #self.add_text('*')

    def depart_literal_emphasis(self, node):
        pass
        dprint()
        #self.add_text('*')

    def visit_strong(self, node):
        pass
        dprint()
        #self.add_text('**')

    def depart_strong(self, node):
        pass
        dprint()
        #self.add_text('**')

    def visit_abbreviation(self, node):
        pass
        dprint()
        #self.add_text('')

    def depart_abbreviation(self, node):
        pass
        dprint()
        #if node.hasattr('explanation'):
        #    self.add_text(' (%s)' % node['explanation'])

    def visit_title_reference(self, node):
        pass
        dprint()
        #self.add_text('*')

    def depart_title_reference(self, node):
        pass
        dprint()
        #self.add_text('*')

    def visit_literal(self, node):
        pass
        dprint()
        #self.add_text('``')

    def depart_literal(self, node):
        pass
        dprint()
        #self.add_text('``')

    def visit_subscript(self, node):
        pass
        dprint()
        #self.add_text('_')

    def depart_subscript(self, node):
        pass
        dprint()
        pass

    def visit_superscript(self, node):
        pass
        dprint()
        #self.add_text('^')

    def depart_superscript(self, node):
        pass
        dprint()
        pass

    def visit_footnote_reference(self, node):
        pass
        dprint()
        #self.add_text('[%s]' % node.astext())
        raise nodes.SkipNode

    def visit_citation_reference(self, node):
        pass
        dprint()
        #self.add_text('[%s]' % node.astext())
        raise nodes.SkipNode

    def visit_Text(self, node):
        dprint()
        self.add_text(node.astext())

    def depart_Text(self, node):
        dprint()
        pass

    def visit_generated(self, node):
        pass
        dprint()
        pass

    def depart_generated(self, node):
        pass
        dprint()
        pass

    def visit_inline(self, node):
        pass
        dprint()
        pass

    def depart_inline(self, node):
        pass
        dprint()
        pass

    def visit_problematic(self, node):
        pass
        dprint()
        #self.add_text('>>')

    def depart_problematic(self, node):
        pass
        dprint()
        #self.add_text('<<')

    def visit_system_message(self, node):
        pass
        dprint()
        #self.new_state(0)
        #self.add_text('<SYSTEM MESSAGE: %s>' % node.astext())
        #self.end_state()
        raise nodes.SkipNode

    def visit_comment(self, node):
        pass
        dprint()
        raise nodes.SkipNode

    def visit_meta(self, node):
        pass
        dprint()
        # only valid for HTML
        raise nodes.SkipNode

    def visit_raw(self, node):
        pass
        dprint()
        #if 'text' in node.get('format', '').split():
        #    self.body.append(node.astext())
        raise nodes.SkipNode

    def unknown_visit(self, node):
        pass
        dprint()
        raise NotImplementedError('Unknown node: ' + node.__class__.__name__)

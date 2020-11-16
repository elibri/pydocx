
from pydocx.export import PyDocXHTMLExporter
from pydocx.export.html import HtmlTag, is_only_whitespace, is_not_empty_and_not_only_whitespace
from itertools import chain

#TODO: pomijane są przypisy końcowe. Jeden przypis końcowy jest w pliku footnotes


#https://pydocx.readthedocs.io/en/latest/extending.html

BLOCK_ELEMENTS = ['document', 'body', 'head', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'ol', 'ul', 'li', 'table', 'tr', 'td', 'footnotes', 'footnote']

#patrz na listę metod na napisania w self.node_type_to_export_func_map w pydocx.export.html linia 36

def buffer_elements_should_be_melted(buffer):
  tags = ''.join(tag.to_html() for tag in buffer if isinstance(tag, HtmlTag))
  return tags == "</em><em>" or tags == "</strong><strong>" or tags == "</underline><underline>" or tags == "</sub><sub>" or tags == "</sup><sup>"

#chcę wyeliminować zagnieżdżenia paragrafów
def tokens_without_nested_paras(stream):
  opened_para = False
  
  for token in stream:
    if isinstance(token, HtmlTag) and token.tag == "p":
      if token.closed:
        if opened_para: #zamykaj tag tylko wtedy, gdy jest otwarty
          yield token 
        opened_para = False
      else: #napotykam otwarcie paragrafu
        if opened_para:
          yield HtmlTag(tag="p", closed=True)
        opened_para = True
        yield token
    else:
      yield token

#tworzę bufor dwuelementowy, który pozwala mi eliminować zagnieżdżenie paragrafów oraz eleminować kombinacje takie jak </em><em>
def tokens_without_reduntant_inline_tags(stream):
  buffer = []

  for token in stream:
  
    if buffer_elements_should_be_melted(buffer):
      buffer = []

    if len(buffer) < 2:
      buffer.append(token)
    else:
      yield buffer[0]
      buffer = buffer[1:] + [token]

  for token in buffer:
    yield token

class DocXMLExporter(PyDocXHTMLExporter):

  def export_document(self, document):
    tag = HtmlTag('document')
    if not 'footnotes' in dir(self): #plik parsowany jest dwukrotnie (parz first_pass w base.py)
      self.footnotes = {}            #przed drugim przejściem nie chcemy stracić wyników
    results = super(PyDocXHTMLExporter, self).export_document(document)
    sequence = []
    #head = self.head()
    #if head is not None:
    #  sequence.append(head)
    if len(self.footnotes) > 0:
      sequence.append(self.export_footnote_texts())
    if results is not None:
      sequence.append(results)

    return tag.apply(chain(*sequence))

  def export_footnote_texts(self):
    yield HtmlTag('footnotes', closed=False)
    for footnote_id, tokens in self.footnotes.items():
      yield HtmlTag('footnote', closed=False, id=footnote_id)
      for token in tokens:
        yield token
      yield HtmlTag('footnote', closed=True)
    yield HtmlTag('footnotes', closed=True)

  def tokens_with_indentations(self):
    level = 0
    eof_emitted = False
    yield '<?xml version="1.0" encoding="UTF-8"?>'
    for token in tokens_without_nested_paras(tokens_without_reduntant_inline_tags(super(PyDocXHTMLExporter, self).export())):
      if isinstance(token, HtmlTag):
        if token.tag in BLOCK_ELEMENTS:
          if token.closed:
            level = level - 1
            if eof_emitted:
              yield " " * level
            yield token.to_html() + "\n"
            eof_emitted = True
          else:
            if not eof_emitted:
              yield "\n"
            yield " " * level + token.to_html()
            eof_emitted = False
            level = level + 1
        else:
          yield token.to_html()
      else:
        yield token
        eof_emitted = False

  def export(self):
    return ''.join(token for token in self.tokens_with_indentations()).strip()

  #to jest kopia export_run_property
  #zamiast aplikować tag, zmieniam tekst na wielkie litery
  def export_uppercased_run_property(self, run, results):
    for result in results:
      if is_only_whitespace(result):
        yield result
      else:
        results = chain([result], results)
        break
    else:
      results = None

    if results:
      for result in results:
        if isinstance(result, HtmlTag):
          yield result
        else:
          yield result.upper()

  def get_hyperlink_tag(self, target_uri):
    pass

  def export_run_property_underline(self, run, results):
    tag = HtmlTag('underline')
    return self.export_run_property(tag, run, results)

  def export_run_property_caps(self, run, results):
    return self.export_uppercased_run_property(run, results)

  def export_run_property_small_caps(self, run, results):
    return self.export_uppercased_run_property(run, results)

  def export_run_property_dstrike(self, run, results):
    tag = HtmlTag('strike')
    return self.export_run_property(tag, run, results)

  def export_run_property_strike(self, run, results):
    tag = HtmlTag('strike')
    return self.export_run_property(tag, run, results)

  def export_run_property_vanish(self, run, results):
    pass

  def export_run_property_hidden(self, run, results):
    pass

  def export_run_property_color(self, run, results):
    return results

  def export_paragraph(self, paragraph):
    results = super(PyDocXHTMLExporter, self).export_paragraph(paragraph)

    results = is_not_empty_and_not_only_whitespace(results)
    if results is None:
      return

    tag = self.get_paragraph_tag(paragraph)
    if tag:
      alignment = paragraph.effective_properties.justification
      if alignment and alignment != "left":
        tag.attrs['align'] = alignment
      results = tag.apply(results)

    for result in results:
      yield result

  def export_tab_char(self, tab_char):
    return "\t"

  def export_paragraph_property_justification(self, paragraph, results):
    return results

  def export_paragraph_property_indentation(self, paragraph, results):
    return results

  def export_run_property_vertical_align_superscript(self, run, results):
    if results is not None:
      results = list(results)
      if len(results) == 1 and isinstance(results[0], HtmlTag) and results[0].tag == "footnotemark":
        yield results[0]
      elif len(results) > 0:
        yield HtmlTag(tag='sup', closed=False)
        for token in results:
          yield token
        yield HtmlTag(tag='sup', closed=True)

  def export_footnote_reference(self, footnote_reference):
    ftokens = chain(*(list(self.node_type_to_export_func_map[type(child)](child)) for child in footnote_reference.footnote.children))
    self.footnotes[footnote_reference.footnote_id] = [token for token in ftokens if token != '\t']
    yield HtmlTag(tag="footnotemark", id=footnote_reference.footnote_id, allow_self_closing=True)

  def export_numbering_span(self, numbering_span):
    results = super(PyDocXHTMLExporter, self).export_numbering_span(numbering_span)
    attrs = {}
    tag_name = 'ul'
    if not numbering_span.numbering_level.is_bullet_format():
      attrs['list-style-type'] = numbering_span.numbering_level.num_format
      tag_name = 'ol'
    tag = HtmlTag(tag_name, **attrs)
    return tag.apply(results)

  def export_footnote_reference_mark(self, footnote_reference_mark):
    pass

  def footer(self):
    return []

  def export_listing_paragraph_property_indentation(self, paragraph, level_properties, include_text_indent=False):
    return {}

def doc2xml(path):
  return DocXMLExporter(path).export()



from pydocx.export.xml import tokens_without_reduntant_inline_tags
import sys, os, re
from pydocx.export.html import HtmlTag

def ot(tagname):
  return HtmlTag(tag=tagname)

def ct(tagname):
  return HtmlTag(tag=tagname, closed=True)

def stringify(stream):
  for token in stream:
    if isinstance(token, HtmlTag):
      yield token.to_html()
    else:
      yield token

def test_removing_siblibgs_em():
  stream = [ot("body"), ot("p"), "To jest ", ot("em"), "napisane ", ct("em"), ot("em"), "kursywą", ct("em"), ct("p"), ct("body")]
  
  assert "<body><p>To jest <em>napisane kursywą</em></p></body>" == ''.join(token for token in stringify(tokens_without_reduntant_inline_tags(stream)))





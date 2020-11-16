
from pydocx.export.xml import doc2xml
import sys, os, re

def strip_heredoc(text):
  indent = len(min(re.findall(r'\n[ \t]*(?=\S)', text) or ['']))
  pattern = r'\n[ \t]{%d}' % (indent - 1)
  return re.sub(pattern, '\n', text).strip()


def fixture_doc2xml(docxname):
  fn = os.path.join(sys.path[0], "tests", "fixtures", docxname)
  return doc2xml(fn)


def test_headers():
  xml = fixture_doc2xml("headers.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <h1>Nagłówek 1</h1>
      <h2>Nagłówek 2</h2>
      <p>Zwykły tekst z <em>kursywą</em> i <strong>pogrubieniem</strong>.</p>
     </body>
    </document>""")

  assert expected == xml

def test_footnotes():
  xml = fixture_doc2xml("footnotes.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <footnotes>
      <footnote id="2">
       <p>Treść pierwszego przypisu</p>
      </footnote>
     </footnotes>
     <body>
      <p>To jest zdanie z dwoma<footnotemark id="2" /> przypisami.</p>
     </body>
    </document>""")

  assert expected == xml


def test_inline_styles():
  xml = fixture_doc2xml("all_configured_styles.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p><strong>aaa</strong></p>
      <p><underline>bbb</underline></p>
      <p><em>ccc</em></p>
      <p>DDD</p>
      <p>EEE</p>
      <p><strike>fff</strike></p>
      <p><strike>ggg</strike></p>
     </body>
    </document>""")

  assert expected == xml

def test_combined_inline_styles():
  xml = fixture_doc2xml("combined_inline_styles.docx")
  
  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p><em><strong>aaa</strong></em></p>
      <p><underline><em>bbb</em></underline></p>
      <p><em><strong>DDD</strong></em></p>
      <p><em><strong>EEE</strong></em></p>
     </body>
    </document>""")
  
  assert expected == xml

def test_stripping_links():
  xml = fixture_doc2xml("inline_tags.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>This sentence has some <strong>bold</strong>, some <em>italics</em> and some <underline>underline</underline>, as well as a hyperlink.</p>
     </body>
    </document>""")

  assert expected == xml

def test_super_and_subscript():
  xml = fixture_doc2xml("super_and_subscript.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>AAA<sup>BBB</sup></p>
      <p><sub>CCC</sub>DDD</p>
     </body>
    </document>""")

  assert expected == xml
 
def test_has_title():
  xml = fixture_doc2xml("has_title.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>Title</p>
      <p>Text</p>
     </body>
    </document>""")

  assert expected == xml


def test_justification():
  xml = fixture_doc2xml("justification.docx")
 
  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p align="center">Center Justified</p>
      <p align="right">Right justified</p>
      <p align="right">Right justified and pushed in from right</p>
      <p align="center">Center justified and pushed in from left and it is great and it is the coolest thing of all time and I like it and I think it is cool</p>
      <p>Left justified and pushed in from left</p>
     </body>
    </document>""")

  assert expected == xml
 
def test_simple_list():
  xml = fixture_doc2xml("simple_lists.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <ol list-style-type="decimal">
       <li>One</li>
      </ol>
      <ul>
       <li>two</li>
      </ul>
     </body>
    </document>""")

  assert expected == xml
 
def test_nested_lists():

  xml = fixture_doc2xml("nested_lists.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <ol list-style-type="decimal">
       <li>one</li>
       <li>two</li>
       <li>three
        <ol list-style-type="decimal">
         <li>AAA</li>
         <li>BBB</li>
         <li>CCC
          <ol list-style-type="decimal">
           <li>alpha</li>
          </ol>
         </li>
        </ol>
       </li>
       <li>four</li>
      </ol>
      <ol list-style-type="decimal">
       <li>xxx
        <ol list-style-type="decimal">
         <li>yyy</li>
        </ol>
       </li>
      </ol>
      <ul>
       <li>www
        <ul>
         <li>zzz</li>
        </ul>
       </li>
      </ul>
     </body>
    </document>""")

  assert expected == xml

def test_styled_color():

  xml = fixture_doc2xml("styled_color.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>AAA</p>
      <p>BBB</p>
      <table border="1">
       <tr>
        <td>CCC</td>
       </tr>
      </table>
      <p>DDD</p>
     </body>
    </document>""")

  assert expected == xml

def test_tabs():
  xml = fixture_doc2xml("include_tabs.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>AAA\tBBB</p>
     </body>
    </document>""")

  assert expected == xml

def test_paragraph_with_margins():

  xml = fixture_doc2xml("paragraph_with_margins.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>Heading1</p>
      <p>Heading2</p>
      <p>\t\tHeading3</p>
     </body>
    </document>""")
  
  assert expected == xml

def test_br():
  xml = fixture_doc2xml("shift_enter.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p>AAA<br />BBB</p>
      <p>CCC</p>
      <ol list-style-type="decimal">
       <li>DDD<br />EEE</li>
       <li>FFF</li>
      </ol>
      <table border="1">
       <tr>
        <td>GGG<br />HHH</td>
        <td>III<br />JJJ</td>
       </tr>
       <tr>
        <td>KKK</td>
        <td>LLL</td>
       </tr>
      </table>
     </body>
    </document>""")

  assert expected == xml

def test_textbox():
  xml = fixture_doc2xml("textbox.docx")

  expected = strip_heredoc("""
    <?xml version="1.0" encoding="UTF-8"?>
    <document>
     <body>
      <p></p>
      <p>AAA</p>
      <p>BBB</p>
      <p>CCCDDD</p>
     </body>
    </document>""")

  assert expected == xml


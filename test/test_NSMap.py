from unittest import TestCase

import main


class TestNSMap(TestCase):
    def test_ns(self):
        test_map = {
             'ooow': 'http://openoffice.org/2004/writer',
             'dr3d': 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
             'script': 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
             'fo': 'urn:oasis:names:tc:opendocument:'
                   'xmlns:xsl-fo-compatible:1.0',
             'calcext': 'urn:org:documentfoundation:names:experimental:'
                        'calc:xmlns:calcext:1.0',
             'form': 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
             'tableooo': 'http://openoffice.org/2009/table',
             'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
             'dom': 'http://www.w3.org/2001/xml-events',
             'of': 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
             'grddl': 'http://www.w3.org/2003/g/data-view#',
             'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
             'rpt': 'http://openoffice.org/2005/report',
             'drawooo': 'http://openoffice.org/2010/draw',
             'xlink': 'http://www.w3.org/1999/xlink',
             'dc': 'http://purl.org/dc/elements/1.1/',
             'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
             'xsd': 'http://www.w3.org/2001/XMLSchema',
             'chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
             'loext': 'urn:org:documentfoundation:names:experimental:'
                      'office:xmlns:loext:1.0',
             'css3t': 'http://www.w3.org/TR/css3-text/',
             'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
             'presentation': 'urn:oasis:names:tc:opendocument:'
                             'xmlns:presentation:1.0',
             'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
             'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
             'xforms': 'http://www.w3.org/2002/xforms',
             'formx': 'urn:openoffice:names:experimental:'
                      'ooxml-odf-interop:xmlns:form:1.0',
             'math': 'http://www.w3.org/1998/Math/MathML',
             'ooo': 'http://openoffice.org/2004/office',
             'field': 'urn:openoffice:names:'
                      'experimental:ooo-ms-interop:xmlns:field:1.0',
             'xhtml': 'http://www.w3.org/1999/xhtml',
             'oooc': 'http://openoffice.org/2004/calc',
             'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
             'number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
             'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
        }

        ns_map = main.NSMap(test_map)

        for key in test_map:
            ending = 'somerandomteststring'
            result = ns_map.ns(key + ':' + ending)
            correct = '{%s}%s' % (test_map[key], ending)

            if result != correct:
                self.fail('result %s did not match expected %s' % (
                    result, correct))

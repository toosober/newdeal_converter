from main import Parser


def test_parse_document():
    document = Parser().parse_document('test_data.xls')
    with open('test_data.txt', 'r') as doc:
        document_str_should_be = doc.read()
    assert document.__str__() == document_str_should_be

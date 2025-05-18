def insert_text_at_top(doc_id, text, docs_service):
    """
    Inserts plain text at the top of the given Google Doc.
    """
    requests = [
        {
            'insertText': {
                'location': {
                    'index': 1  # index 1 = after the start of the doc
                },
                'text': text + '\n'
            }
        }
    ]

    response = docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={'requests': requests}
    ).execute()

    print(f"Inserted text into document: https://docs.google.com/document/d/{doc_id}")


if __name__ == '__main__':
    docs_service = authenticate_gdoc(SERVICE_ACCOUNT_FILE, SCOPES)
    insert_text_at_top('YOUR_TEST_DOC_ID', 'Hello, this was added by Python!', docs_service)

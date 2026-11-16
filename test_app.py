import io
import json
import tempfile
import threading
import unittest
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from openpyxl import load_workbook
import server

class ContactBookTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        server.DB = Path(cls.tmp.name)/'test.db'
        server.TOKEN = 'test-secret'
        cls.http = server.ThreadingHTTPServer(('127.0.0.1', 0), server.Handler)
        cls.url = 'http://127.0.0.1:' + str(cls.http.server_port)
        threading.Thread(target=cls.http.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown(); cls.http.server_close(); cls.tmp.cleanup()

    def req(self, path, method='GET', data=None, auth=True):
        headers = {'Authorization': 'Bearer test-secret'} if auth else {}
        if isinstance(data, dict): data = json.dumps(data).encode()
        return urlopen(Request(self.url+path, data=data, method=method, headers=headers))

    def test_full_workflow(self):
        with self.assertRaises(HTTPError) as e: self.req('/api/contacts', auth=False)
        self.assertEqual(e.exception.code, 401)
        item = {'name': '=test 中文', 'favorite': True, 'notes': '备注', 'methods': [{'type':t,'value':v} for t,v in [('phone','00123'),('phone','456'),('email','test@example.com'),('social','微信'),('address','福州')]]}
        created = json.load(self.req('/api/contacts','POST',item)); ident=created['id']
        self.assertEqual(json.load(self.req('/api/contacts'))[0]['name'], item['name'])
        data = self.req('/api/export').read()
        wb = load_workbook(io.BytesIO(data)); self.assertEqual(wb.active['A2'].data_type,'s')
        self.assertEqual(server.import_xlsx(data)[0],item)
        self.assertEqual(json.load(self.req('/api/import','POST',data))['imported'],1)
        item['favorite']=False
        self.req('/api/contacts/'+str(ident),'PUT',item).close()
        self.assertEqual(len(json.load(self.req('/api/contacts'))),2)
        self.req('/api/contacts/'+str(ident),'DELETE').close()
        self.assertEqual(len(json.load(self.req('/api/contacts'))),1)
        before=json.load(self.req('/api/contacts'))
        with self.assertRaises(HTTPError): self.req('/api/import','POST',b'invalid workbook')
        self.assertEqual(json.load(self.req('/api/contacts')),before)
        with self.assertRaises(HTTPError): self.req('/api/contacts','POST',{'name':''})
        with self.assertRaises(HTTPError): self.req('/api/contacts/999','DELETE')

if __name__=='__main__': unittest.main()

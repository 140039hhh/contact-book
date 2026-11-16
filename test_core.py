"""Exercise request handlers without requiring local network permissions."""
import io
import json
import tempfile
import unittest
from pathlib import Path
import server

class CoreTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        server.DB = Path(self.temp.name)/'contacts.db'
        server.TOKEN = 'secret'

    def tearDown(self): self.temp.cleanup()

    def request(self, path, method='GET', data=b'', auth=True):
        handler = object.__new__(server.Handler)
        if isinstance(data, dict): data=json.dumps(data).encode()
        handler.path=path; handler.command=method
        handler.headers={'Content-Length':str(len(data))}
        if auth: handler.headers['Authorization']='Bearer secret'
        handler.rfile=io.BytesIO(data)
        result=[]
        handler.reply=lambda *args: result.append(args)
        handler.handle_request()
        return result[0]

    def test_crud_auth_and_excel(self):
        self.assertEqual(self.request('/api/contacts',auth=False)[0],401)
        item=dict(name='=中文联系人',favorite=True,notes='备注',methods=[dict(type=t,value=v) for t,v in [('phone','00123'),('phone','456'),('email','a@example.com'),('social','微信'),('address','福州')]])
        code,c=self.request('/api/contacts','POST',item);self.assertEqual(code,201)
        data=self.request('/api/export')[1]
        self.assertEqual(server.import_xlsx(data),[item])
        self.assertEqual(self.request('/api/import','POST',data)[1]['imported'],1)
        item['favorite']=False
        self.assertEqual(self.request('/api/contacts/'+str(c['id']),'PUT',item)[0],200)
        self.assertFalse(self.request('/api/contacts')[1][1]['favorite'])
        self.assertEqual(self.request('/api/contacts/'+str(c['id']),'DELETE')[0],200)
        self.assertEqual(len(self.request('/api/contacts')[1]),1)
        self.assertEqual(self.request('/api/import','POST',b'broken')[0],400)
        self.assertEqual(len(self.request('/api/contacts')[1]),1)
        self.assertEqual(self.request('/api/contacts','POST',{'name':''})[0],400)
        self.assertEqual(self.request('/api/contacts/999','DELETE')[0],404)

    def test_invalid_import_is_atomic(self):
        from openpyxl import Workbook
        wb=Workbook();ws=wb.active;ws.append(['name','favorite','notes']);ws.append(['valid',0,'']);ws.append(['',0,'invalid'])
        buf=io.BytesIO();wb.save(buf)
        self.assertEqual(self.request('/api/import','POST',buf.getvalue())[0],400)
        self.assertEqual(self.request('/api/contacts')[1],[])

    def test_excel_without_dimensions(self):
        from openpyxl import Workbook
        wb=Workbook(write_only=True);ws=wb.create_sheet()
        ws.append(['name','favorite','notes','phone_1'])
        ws.append(['流式Excel测试',1,'无行列元数据','00123'])
        buf=io.BytesIO();wb.save(buf)
        result=self.request('/api/import','POST',buf.getvalue())
        self.assertEqual(result[0],200)
        contact=self.request('/api/contacts')[1][0]
        self.assertTrue(contact['favorite'])
        self.assertEqual(contact['methods'][0]['value'],'00123')

if __name__=='__main__':unittest.main()

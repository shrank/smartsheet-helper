# flake8: noqa: E510,E722
import smartsheet
import os
import copy
from datetime import datetime

def create_multivalue(values):
    res = smartsheet.models.MultiPicklistObjectValue()
    res.values = values
    return res

class smartsheet_helper:
    def __init__(self):
        self.sheet = os.getenv("SMARTSHEET_SHEET_ID")
        self.smart = smartsheet.Smartsheet()
        self.smart.errors_as_exceptions(True)
        self.data = smartsheet.models.Sheet()
        self.columns = {}
        self.contact_columns={}
        self.last_timestamp = None
        self.updateRows={}

    def loadColumns(self):
        self.columns = {}
        self.data.columns = self.smart.Sheets.get_columns(self.sheet).data
        for a in self.data.columns:
            self.columns[a.title] = a.id
        
    def get_copy(self):
        data = copy.copy(self)
        data.data = None
        data.columns = {}
        data.last_timestamp = None
        data.updateRows={}
        return data

    def getAll(self):
        self.data = self.smart.Sheets.get_sheet(self.sheet)
        self.last_timestamp = datetime.utcnow().isoformat()
        print(self.last_timestamp)
        for a in self.data.columns:
            self.columns[a.title] = a.id
            if(a.type == "CONTACT_LIST"):
                self.contact_columns[a.title] = {"type": "CONTACT_LIST", "contactOptions": [] }
        return self.data.rows

    def getUpdated(self):
        if(self.last_timestamp is None):
            return self.getAll()
        self.data = self.smart.Sheets.get_sheet(self.sheet, rows_modified_since=self.last_timestamp)
        self.last_timestamp=datetime.utcnow().isoformat()
        for a in self.data.columns:
            self.columns[a.title] = a.id
        return self.data.rows

    def dict2row(self,values, row=None, diff=False, skip_nonexistend=True):
        new_row = smartsheet.models.Row()
        changed = False
        if (row is not None):
            new_row.id = row.id
        for k, v in values.items():
            if(skip_nonexistend and k not in self.columns):
                # skip nonextistent cells
                continue
            cell = smartsheet.models.Cell()
            if (row is not None):
                cell = self.getCell(row, k)
                if (diff is True):
                    if(cell.value == v):
                        continue
            changed = True
            if(isinstance(v, smartsheet.models.MultiPicklistObjectValue)):
                cell._value = None
                cell.object_value = v
            else:
                cell.value = v
            # try if(isinstance(v, smartsheet.models.Contact)):
            if(k in self.contact_columns):
                print(k, v)
                self.addContact(k,v)
                cell.value = {"email": "", "name":v}
            cell.column_id = self.columns[k]
            new_row.cells.append(cell)
        if (changed is True):
            return new_row
        return None

    def update(self, row, values, diff=False):
        self.addUpdate(row, values, diff)
        return self.commitUpdate()

    def addUpdate(self, row, values, diff=False):
        new_row = self.dict2row(values, row, diff)
        if(new_row != None):
            self.updateRows[new_row.id] = new_row

    def commitUpdate(self):
        if (len(self.updateRows) > 0):
            rows=[]
            for a in self.updateRows:
                rows.append(self.updateRows[a])
            res = self.smart.Sheets.update_rows(self.sheet, rows)
            self.updateRows={}
            return res

    def addContact(self, column, value):
        if(column not in self.contact_columns):
            return
        self.contact_columns[column]["contactOptions"].append({"email": "", "name": value})
        self.smart.Sheets.update_column(
            self.sheet, 
            self.columns[column], 
            self.contact_columns[column])

    def insert(self, values):
        new_row = self.dict2row(values)
        new_row.to_top = True
        return self.smart.Sheets.add_rows(self.sheet, [new_row])

    def insert_bulk(self, rows):
        res = []
        for values in rows:
          a = self.dict2row(values)
          a.to_top = True
          res.append(a)
        return self.smart.Sheets.add_rows(self.sheet, res)

    def getValue(self, row, key, default=None):
        res = self.getCell(row, key)
        if(res is None or res.value is None):
            return default
        return res.value

    def getCell(self, row, key):
        if(key not in self.columns.keys()):
            return None
        for a in row.cells:
            if(a.column_id == self.columns[key]):
                return a
        return None

    def find_first_row(self, column, value):
        for a in self.data.rows:
            if(self.getCell(a,column).value == value):
                return a
        return None
    
    def attachment_to_row(self,row_id, file_data, name):
        return self.smart.Attachments.attach_file_to_row(self.sheet, row_id, (name, file_data))
    
    def comments_to_row(self, row, comments):
        id = None
        for c in comments:
            n = smartsheet.models.comment.Comment()
            n.text = c

            if(id == None):

              if(len(row.discussions) > 0):
                id = row.discussions[0].id
              else:
                a = self.smart.Discussions.get_row_discussions(self.sheet, row.id).data
                if(len(a) > 0):
                    id = a[0].id
                else:
                    a = self.smart.Discussions.create_discussion_on_row(self.sheet, row.id,  smartsheet.models.discussion.Discussion(props={"comment":n})).data
                    id = a.id
                    continue
              if(id == None):
                  return

            self.smart.Discussions.add_comment_to_discussion(self.sheet, id, n)

    def get_webhooks(self):
        return self.smart.Webhooks.list_webhooks().data

    def create_webhook(self,name, url, columns=None, enabled=True):
        item = smartsheet.models.webhook.Webhook()
        item.callback_url = url
        item.name = name
        item.events =  '*.*'
        item.scope = "sheet"
        item.version = 1
        item.scope_object_id = int(self.sheet)
        if(columns is not None):
            if(len(self.columns) == 0):
                self.loadColumns()
            subscope=[]
            for a in columns:
                if(a in self.columns):
                    subscope.append(self.columns[a])
            item.subscope = smartsheet.models.webhook_subscope.WebhookSubscope(props={"column_ids":subscope})
        res = self.smart.Webhooks.create_webhook(item).data
        if(res.enabled != enabled):
            self.smart.Webhooks.update_webhook(res.id_, smartsheet.models.webhook.Webhook(props={"enabled": enabled}))

        

    def delete_webhook(self, name):
        for a in self.get_webhooks(self):
            if(a.name == name):
                self.smart.Webhooks.delete_webhook(a.id_)

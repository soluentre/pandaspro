from pandaspro.email.utlis import replace_with_dict
from jinja2 import Template
import win32com.client as win32


class DataFetcher:
    """
    This is the base class for specific data checking classes
    For example, fetch staff on board info for STA/DAIS Assignment email
    """
    def fetch_data(
            self,
            **kwargs
    ):
        raise NotImplementedError('Subclasses must implement this method')

    def fetch_showitems(
            self,
            **kwargs
    ):
        raise NotImplementedError('Subclasses must implement this method')


def create_mail_class(template_path, data_fetcher_class):
    class Mail:
        def __init__(self, **kwargs):
            self.template = template_path
            self.fetcher = data_fetcher_class()
            self.data = self.fetcher.fetch_data(**kwargs)
            self.showitems = self.fetcher.fetch_showitems(**kwargs)  # this is corresponding to the design in email template (show which part)

            # Render the email with data automatically in init
            with open(self.template, 'r', encoding='utf-8') as file:
                self.html_raw = file.read()
            html_final = Template(self.html_raw).render(self.showitems)
            self.html_final = replace_with_dict(html_final, self.data)

            # Create this email object
            self.olApp = win32.Dispatch('Outlook.Application')
            self.olNS = self.olApp.GetNameSpace('MAPI')
            self.mail = self.olApp.CreateItem(0)
            self.mail.To = self.data['__to__'] if '__to__' in self.data.keys() else ''
            self.mail.CC = self.data['__cc__'] if '__cc__' in self.data.keys() else ''
            self.mail.BCC = self.data['__bcc__'] if '__bcc__' in self.data.keys() else ''
            self.mail.Subject = self.data['__subject__'] if '__subject__' in self.data.keys() else ''
            self.mail.HTMLBody = self.html_final

        def attach(self, attachments: list):
            for item in attachments:
                self.mail.Attachments.Add(item)
            return self.mail

        def display(self):
            self.mail.Display()
            return self.mail

    return Mail



if __name__ == '__main__':
    # att = [
    #     r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Emails and Manuals\STA and DAIS\T200011_Short_Term_Assignment_Developmental_Assignment_Memorandum.pdf',
    # ]
    # mytemplate = email_template(
    #     'sta_dais_init',
    #     assign_type='dais',
    #     rec_mgr_upi=300600,
    #     staff_upi=607313,
    #     position=71058
    # )
    # e = email(mytemplate)
    # e.attach(att).display
    pass

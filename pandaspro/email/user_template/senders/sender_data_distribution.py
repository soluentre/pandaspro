from pandaspro.email.email import create_mail_class
from pandaspro.email.user_template.engines.data_distribution import data_distribution

data_distribution_Sender = create_mail_class(
    r'pandaspro/email/user_template/email_template/data_distribution.html',
    data_distribution
)
myemail = data_distribution_Sender(ifscode=111)
show = myemail.display()
del create_mail_class
del data_distribution
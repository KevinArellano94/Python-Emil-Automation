import win32com.client as client
from datetime import date

def main():
    e_( n )
    outlook = client.Dispatch("Outlook.Application")

    e_( n )
    message = outlook.CreateItem(0)
    
    e_( n )
    message.Display()

    e_( n )
    message.To, message.CC, message.BCC = get_email_recipient()

    e_( n )
    today = date.today()
    d1 = today.strftime("%m.%d.%Y")
    message.Subject = "Email Subject | " + d1

    e_( n )
    message.HTMLBody = get_body()
    
    e_( n )
    # message.Save()

    e_( n )
    # message.Send()

    e_( n )

def get_email_recipient():
    email_to = "Kevin.Arellano94@EMail.Com"
    email_cc = "Kevin.Arellano94@FMail.Com"
    email_bcc = "Kevin.Arellano94@GMail.Com"
    return email_to, email_cc, email_bcc

def get_body():
    html = """
    <div>
        {{ PLACE CONFLUENCE CODE HERE }}
    </div>
    """
    return html

def e_( data ):
    global n
    if n == 0:
        print( str( n ) + " - Launching Outlook email." )
    elif n == 1:
        print( str( n ) + " - Creating email." )
    elif n == 2:
        print( str( n ) + " - Displaying email." )
    elif n == 3:
        print( str( n ) + " - Inlcluded recepient 'To', 'CC' and 'BCC'." )
    elif n == 4:
        print( str( n ) + " - Added Subject with xx.xx.xxxx date variable." )
    elif n == 5:
        print( str( n ) + " - Added email Body." )
    elif n == 6:
        print( str( n ) + " - Saved." )
    elif n == 7:
        print( str( n ) + " - Sent." )
    elif n == 8:
        print( str( n ) + " - Completed." )
    n += 1

if __name__ == "__main__":
    n = 0
    main()

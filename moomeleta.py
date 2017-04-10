import httplib2
import os
import oauth2client
from oauth2client import client, tools
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors, discovery
import mimetypes
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from xlr import open_file

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Send Email'

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'gmail-aa.json')
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials

def SendMessage(sender, to, subject, msgHtml, msgPlain, attachmentFile=None):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    if attachmentFile:
        message1 = createMessageWithAttachment(sender, to, subject, msgHtml, msgPlain, attachmentFile)
    else:
        message1 = CreateMessageHtml(sender, to, subject, msgHtml, msgPlain)
    result = SendMessageInternal(service, "me", message1)
    return result

def SendMessageInternal(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print 'Message Id: %s' % message['id']
        return message
    except errors.HttpError, error:
        print 'An error occurred: %s' % error
        return "Error"
    return "OK"

def CreateMessageHtml(sender, to, subject, msgHtml, msgPlain):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to
    msg.attach(MIMEText(msgPlain, 'plain'))
    msg.attach(MIMEText(msgHtml, 'html'))
    return {'raw': base64.urlsafe_b64encode(msg.as_string())}

def createMessageWithAttachment(
    sender, to, subject, msgHtml, msgPlain, attachmentFile):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      msgHtml: Html message to be sent
      msgPlain: Alternative plain text message for older email clients
      attachmentFile: The path to the file to be attached.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart('mixed')
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    messageA = MIMEMultipart('alternative')
    messageR = MIMEMultipart('related')

    messageR.attach(MIMEText(msgHtml, 'html'))
    messageA.attach(MIMEText(msgPlain, 'plain'))
    messageA.attach(messageR)

    message.attach(messageA)

    print "create_message_with_attachment: file:", attachmentFile
    content_type, encoding = mimetypes.guess_type(attachmentFile)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(attachmentFile, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(attachmentFile, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(attachmentFile, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(attachmentFile, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(attachmentFile)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_string())}

def main(name, to):

    chunk = """

        <div dir="ltr">
            <p style="background-image:initial;background-position:initial;background-size:initial;background-repeat:initial;background-origin:initial;background-clip:initial"><span style="font-size:9.5pt;font-family:tahoma,sans-serif">Dear %s,</span><span style="font-size:9.5pt;font-family:arial,sans-serif"><span></span></span>
            </p>

            <p style="text-align:justify;background-image:initial;background-position:initial;background-size:initial;background-repeat:initial;background-origin:initial;background-clip:initial"><span style="font-size:9.5pt;font-family:tahoma,sans-serif">I take pleasure in introducing&nbsp;<b>National Institute of Technology (NIT) Srinagar</b>&nbsp;to your esteemed organization. Now, more than ever, the emphasis is on Institute-Industry Interaction, and both the Institute, conducting the&nbsp;Campus&nbsp;Recruitment programme, and the Industry, participating in the same, are bound to find it mutually beneficial.</span><span style="font-size:9.5pt;font-family:arial,sans-serif"><span></span></span>
            </p>

            <p style="text-align:justify;background-image:initial;background-position:initial;background-size:initial;background-repeat:initial;background-origin:initial;background-clip:initial"><span style="font-size:9.5pt;font-family:tahoma,sans-serif">National Institute of Technology Srinagar is an Institute with total commitment to quality and excellence in academic pursuits. This institute was one of the first eight Regional Engineering Colleges, established in 1960, by the Government of India. National Institute of Technology Srinagar&nbsp;offers B.Tech in eight engineering disciplines namely,&nbsp;<b>Electrical Engineering, Computer Science Engineering, Information Technology, Electronics &amp; Communication Engineering, Mechanical Engineering, Chemical Engineering, Civil Engineering&nbsp;</b>and<b>&nbsp;Metallurgical &amp; Materials Engineering</b>. Further, several M.Tech programs namely,&nbsp;<b>Water Resources Engineering (Civil), Structural Engineering (Civil), Transportation Engineering (Civil), Geo-technical Engineering (Civil), Mechanical System Design (Mechanical)&nbsp;</b>and<b>&nbsp;Communication &amp; IT (Electronics &amp; Communication)</b>&nbsp;are also being offered.</span><span style="font-size:9.5pt;font-family:arial,sans-serif"><span></span></span>
            </p>

            <p style="text-align:justify;background-image:initial;background-position:initial;background-size:initial;background-repeat:initial;background-origin:initial;background-clip:initial"><span style="font-size:9.5pt;font-family:tahoma,sans-serif">We would also like to put forward the possibility of&nbsp;interview through video conferencing.&nbsp;</span><span style="font-size:9.5pt;font-family:arial,sans-serif"><span></span></span>
            </p>

            <p style="text-align:justify;background-image:initial;background-position:initial;background-size:initial;background-repeat:initial;background-origin:initial;background-clip:initial"><span style="font-size:9.5pt;font-family:tahoma,sans-serif">Placement Brochure is attached with this mail for your kind reference.&nbsp; Kindly go through it for being familiar with our institute.</span><span style="font-size:9.5pt;font-family:arial,sans-serif"><span></span></span>
            </p>

            <p style="text-align:justify;background-image:initial;background-position:initial;background-size:initial;background-repeat:initial;background-origin:initial;background-clip:initial"><span style="font-size:9.5pt;font-family:tahoma,sans-serif">For further assistance, you can contact our&nbsp;<b>placement coordinator Mr. Lakhan Pareek (+91-7597707004)</b></span><span style="font-size:9.5pt;font-family:arial,sans-serif"><span></span></span>
            </p>
        </div>


    """ % name

    signature_JN = """

    <div dir="ltr"><p style="margin:0cm 0cm 0.0001pt;background-image:initial;background-position:initial;background-repeat:initial"><br></p><div><b><font color="#0000ff" face="tahoma,sans-serif">Regards</font></b></div><div><font color="#0000ff"><b>Dr. Javed Ahmed Naqash</b></font></div><div><font color="#999999">Head, Department of Training and Placement</font></div><div><div style="font-size:12.8px"><b><img width="90" height="96" src="https://ci6.googleusercontent.com/proxy/MoejYKheSfjVeYgvXtFoIZIaSBc_wqsY2T-SumXf1HWP2KRAtjGSkcR56AROcBFqGGerDscrNCHk0WC-bBcPEwtceryK-QezXQmAhoQ8-aWxqIsFzSCaG9rG1uciRQGV0jyXLeQCD8lR4NToRB9mHgHoeZqiDStWN0_qgJAKn4pg-HNCAAAVO9ICzp4VcdTClLISj11jtjdyZIc=s0-d-e1-ft#https://docs.google.com/uc?export=download&amp;id=0B2iCMH4NgXVyN2VWd3A1MjhiUVk&amp;revid=0B2iCMH4NgXVyckZZamtvdElpb1ovbWVXeGJ1T25pWHdESVpFPQ" class="CToWUd"></b></div><font color="#999999">National Institute of Technology<br>Srinagar -190006</font><br>Phone No.&nbsp;<font face="comic sans ms,sans-serif">0194-2424809</font><br>Mobile No.&nbsp;<font face="comic sans ms,sans-serif">+91-94195 91569, 9906965736</font></div><div><font face="comic sans ms">Email:&nbsp;<i><a href="mailto:placements@nitsri.net" target="_blank">placements@nitsri.net</a>&nbsp;<a href="mailto:javednaqash@nitsri.net" target="_blank">j<wbr>avednaqash@nitsri.net</a></i></font><br>Website:&nbsp;<i><font color="#0000ff"><a href="http://www.nitsri.net/" target="_blank" data-saferedirecturl="https://www.google.com/url?hl=en&amp;q=http://www.nitsri.net/&amp;source=gmail&amp;ust=1491945692871000&amp;usg=AFQjCNEeK6KjiYnRnEtt7YymKci3sddRVg">http://www.nitsri.net</a></font></i></div></div>


    """
    sender = "placements@nitsri.net"
    subject = "Invitation for Placement Drive, NIT Srinagar (July 2017 passing out batch)"
    msgHtml = chunk + signature_JN
    msgPlain = chunk + signature_JN
    #SendMessage(sender, to, subject, msgHtml, msgPlain)
    # Send message with attachment:
    SendMessage(sender, to, subject, msgHtml, msgPlain, 'PLACEMENT BROCHURE.pdf')

if __name__ == '__main__':
    path = "test.xls"
    janlog = open_file(path)
    for naam, pata in janlog:
        main(naam, "placements@nitsri.net")

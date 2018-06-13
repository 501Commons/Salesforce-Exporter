"""Export Module for Salesforce"""
def main():
    """Main entry point"""

    import sys
    from os.path import join

    salesforce_type = str(sys.argv[1])
    client_type = str(sys.argv[2])
    client_emaillist = str(sys.argv[3])

    if len(sys.argv) < 4:
        print ("Calling error - missing inputs.  Expecting " +
               "salesforce_type client_type client_emaillist [Exporter_root]\n")
        return

    if len(sys.argv) == 5:
        Exporter_root = str(sys.argv[4])
    else:
        Exporter_root = "C:\\repo\\Salesforce-Exporter-Private\\Clients\\" + sys.argv[2] + "\\Salesforce-Exporter"

    sys.stdout = open(join(Exporter_root, '..\\Exporter.log'), 'w')
    print('Exporter Startup')

    exporter_directory = join(Exporter_root, "Clients\\" + client_type)
    print "Setting Exporter Directory: " + exporter_directory

    # Export Data
    print "\n\nExporter - Export Data Process\n\n"
    process_data(exporter_directory, salesforce_type, client_type, client_emaillist)

    print "Exporter process completed\n"

def process_data(exporter_directory, salesforce_type, client_type, client_emaillist):
    """Process Data based on data_mode"""

    from os import makedirs
    from os.path import exists

    sendto = client_emaillist.split(";")
    user = 'db.powerbi@501commons.org'
    smtpsrv = "smtp.office365.com"
    subject = "Export Data Results -"
    file_path = exporter_directory + "\\Status"
    if not exists(file_path):
        makedirs(file_path)
    export_path = exporter_directory + "\\Export"
    if not exists(export_path):
        makedirs(export_path)

    body = "Export Data\n\n"

    status_export = ""
    
    # Export data from Salesforce
    try:
        if not "Error" in subject:
            status_export = export_dataloader(exporter_directory,
                                              client_type, salesforce_type)
        else:
            status_export = "Error detected so skipped"
    except Exception as ex:
        subject += " Error Export"
        body += "\n\nUnexpected export error:" + str(ex)
    else:
        body += "\n\nExport\n" + status_export

    if not "Error" in subject:
        subject += " Successful"

    # Send email results
    send_email(user, sendto, subject, body, file_path, smtpsrv)

    return status_export

def contains_data(file_name):
    """Check if file contains data after header"""

    line_index = 1
    with open(file_name) as file_open:
        for line in file_open:
            # Check if line empty
            line_check = line.replace(",", "")
            line_check = line_check.replace('"', '')
            if (line_index == 2 and line_check != "\n"):
                return True
            elif line_index > 2:
                return True

            line_index += 1

    return False

def export_dataloader(exporter_directory, client_type, salesforce_type):
    """Export out of Salesforce using DataLoader"""

    import os
    from os import listdir
    from os.path import join
    from subprocess import Popen, PIPE

    bat_path = exporter_directory + "\\DataLoader"

    return_code = ""
    return_stdout = ""
    return_stderr = ""

    for file_name in listdir(bat_path):
        if not ".sdl" in file_name:
            continue

        # Check if associated csv has any data
        export_name = os.path.splitext(file_name)[0]
        bat_file = (join(bat_path, "RunDataLoader.bat")
                    + " " + salesforce_type + " "  + client_type + " " + export_name)

        message = "Starting Export Process: " + bat_file
        print message
        return_stdout += message + "\n"
        export_process = Popen(bat_file, stdout=PIPE, stderr=PIPE)

        stdout, stderr = export_process.communicate()

        return_code += "\n\nexport_dataloader (returncode): " + str(export_process.returncode)
        return_stdout += "\n\nexport_dataloader (stdout):\n" + stdout
        return_stderr += "\n\nexport_dataloader (stderr):\n" + stderr

        if (export_process.returncode != 0
                or "Error" in return_stdout
                or "We couldn't find the Java Runtime Environment (JRE)" in return_stdout):
            raise Exception("Invalid Return Code", return_code + return_stdout + return_stderr)

        status_path = exporter_directory + "\\status"

        for file_name_status in listdir(status_path):
            file_name_status_full = join(status_path, file_name_status)
            if "error" in file_name_status_full and contains_data(file_name_status_full):
                raise Exception("error file contains data: " + file_name_status_full, (
                    return_code + return_stdout + return_stderr))

    return return_code + return_stdout + return_stderr

def send_email(send_from, send_to, subject, text, file_path, server):
    """Send email via O365"""

    #https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
    import base64
    import os
    import smtplib
    from os.path import basename
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import COMMASPACE, formatdate

    msg = MIMEMultipart()

    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    from os import listdir, remove
    from os.path import isfile, join
    onlyfiles = [join(file_path, f) for f in listdir(file_path)
                 if isfile(join(file_path, f))]

    for file_name in onlyfiles:
        if contains_data(file_name):
            with open(file_name, "rb") as file_name_open:
                part = MIMEApplication(
                    file_name_open.read(),
                    Name=basename(file_name)
                    )

            # After the file is closed
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file_name)
            msg.attach(part)

    server = smtplib.SMTP(server, 587)
    server.starttls()
    server_password = os.environ['SERVER_EMAIL_PASSWORD']
    server.login(send_from, base64.b64decode(server_password))
    text = msg.as_string()
    server.sendmail(send_from, send_to, text)
    server.quit()

    # Delete all status files
    for file_name in onlyfiles:
        try:
            remove(file_name)
        except:
            continue

def send_salesforce():
    """Send results to Salesforce to handle notifications"""
    #Future update to send to salesforce to handle notifications instead of send_email
    #https://developer.salesforce.com/blogs/developer-relations/2014/01/python-and-the-force-com-rest-api-simple-simple-salesforce-example.html

if __name__ == "__main__":
    main()

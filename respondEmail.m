%% SETUP SEND MAIL
% Set preferences of the email being used by Matlab
setpref('Internet','E_mail','tec.device01@gmail.com');
setpref('Internet','SMTP_Server','smtp.gmail.com');
setpref('Internet', 'SMTP_Username', 'tec.device01@gmail.com'); 
setpref('Internet', 'SMTP_Password', '0dY2!c4szbBQ');

props = java.lang.System.getProperties;
props.setProperty('mail.smtp.auth','true');
props.setProperty('mail.smtp.starttls.enable','true');

%% READMAIL
% A simple script highlighting how you can connect to Outlook and
% import emails, including their subjects, bodies & attachements

%% Connecting to Outlook

outlook = actxserver('Outlook.Application');
mapi=outlook.GetNamespace('mapi');
INBOX=mapi.GetDefaultFolder(6);

%% Retrieving last email

count = INBOX.Items.Count; %index of the most recent email.
firstemail=INBOX.Items.Item(count); %imports the most recent email
% secondmail=INBOX.Items.Item(count-1); %imports the 2nd most recent email
subject = firstemail.get('Subject');
body = firstemail.get('Body');

if subject == "testing funcion"
    sendmail('tec.remote01@gmail.com','Test mail with attachment','C:\Users\cluck\Downloads');
end

%% Saving attachments to current directory
%attachments = firstemail.get('Attachments');
%if attachments.Count >=1
%    fname = attachments.Item(1).Filename;
%    dir = pwd;
%    full = [pwd,'\',fname];
%    attachments.Item(1).SaveAsFile(full)
%end
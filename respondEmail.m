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

txt = sprintf('Checking credentials...')
while (1)

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

C = strsplit(body,'\n');
%customer_pin = strsplit(mat2str(C(1,2)),':')



if (subject == "Pin number" && C(1,2) == "Pin:1234")
    
    sendmail('ebldevice2@gmail.com','Device Data','Data for Load, PV Power and Battery State of Charge.','C:\Users\cluck\Documents\MATLAB\EBL\testdata.xlsx');
    delete(firstemail);
    txt2 = sprintf('Credentials found, email sent')
    mdl = 'ECE552ProjectSim_Updated_v2';
    sim(mdl)
    % delete email
    break
end

outlook.release;
% make sure server hang ending

end


%% Saving attachments to current directory
%attachments = firstemail.get('Attachments');
%if attachments.Count >=1
%    fname = attachments.Item(1).Filename;
%    dir = pwd;
%    full = [pwd,'\',fname];
%    attachments.Item(1).SaveAsFile(full)
%end
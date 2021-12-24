import smtplib




def mail_sent(to_,from_,sub_,msg_,pwd_):
    #---session creation
    ob = smtplib.SMTP("smtp.gmail.com",587)
     #---transport layer connection
    try:
        ob.starttls()
        #----login
        ob.login(from_,pwd_)
        #---sending mail
        msg = "Subject: {}\n\n{}".format(sub_,msg_)
        ob.sendmail(from_,to_,msg)
        #---status
        st = ob.ehlo()
        if st[0] == 250:
            return "success"
        else:
            return "failed"
    except smtplib.SMTPAuthenticationError:
        return "failed"
    except:
        return "failed"
    
    #---close connection
    ob.close()
    
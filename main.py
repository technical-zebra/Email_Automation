import win32com.client as win32
import csv
import time

companies = []
emails = []

with open('list.csv') as csv_file:
    csv_reader = csv.reader(csv_file)
    for row in csv_reader:
        print(row)
        companies.append(row[0].strip())
        emails.append(row[1].strip())

# for i in range(len(companies)):
#     print(f"{companies[i]}       {emails[i]}")

for x in range(len(companies)):
    outlook = win32.Dispatch('Outlook.Application')
    message = outlook.CreateItem(0)

    company_name = companies[x]
    print(company_name)

    email_to = emails[x]
    print(email_to)

    message.To = email_to
    message.Subject = "James Cook University Singapore 27th Convergence Conference"
    #display: block;
    message.HTMLBody = """
        <!doctype html>
        <html>
        <head>
            <style>
                body {
                    line-height: 400%;
                }
                img {
                    width: 99%;
                }
            </style>
        </head>
        
        
        <body>
            <img src="https://cdn.discordapp.com/attachments/1028583684909047808/1050058069305999481/316821900_112417125027644_1578679093283271436_n.jpg">
            <p>From: JCU Singapore 27th Convergence Conference Organizing Team</p> 
            <p>To: Company Name</p>
            <p>Date: 08/12/2022</p>
            <br/>
            <p><b>Ref: James Cook University (JCU), Singapore’s 27th Convergence Conference</b></p>
            <hr>
            <p>Warmest Greetings,</p>
            <br/>
            <p>We take pleasure to invite you as one of the sponsors for the upcoming 27th Convergence Conference of James Cook
                University this 16 January 2023.</p>
            <br/>
            <p>As event organizers for James Cook University’s 27th Convergence Conference, we take pride in our university’s
                sound
                reputation and consistent ranking in the top 400 academic universities worldwide since 2010. In 2016, JCU ranked
                in
                the top two percent of universities in the world by ARWU (Academic Ranking of World Universities).</p>
            <br/>
            <p>The 27th Convergence Conference is the culmination of our graduating students’ project. Close to 300 students
                will be
                presenting their research and case studies of multi-disciplinary topics and research based initiatives ranging
                from
                Marketing, Information Technology (IT), Tourism & Business.
                There are close to 3,000 students in the Singapore campus. VIPs, academics and speakers will also be invited.
                The
                best projects will also be awarded prizes in recognition of their relevance and excellence.</p>
            <br/>
            <p>We sincerely hope we can work together and promote your brand and products. Thank you and we are looking forward
                to work
                with you and your organization.</p>
            <br/>
            <p>PS: refer to the Sponsorship Tier as attached below.</p>
            <br/>
            <p>Best regards,</p>
            <hr>
            <p>Shi Yingjie</p>
            <p>Chairman</p>
            <p>(+65) 90547665 or Jimmyshiyingjie@gmail.com</p>
            <br/>
            <p>Pham Hoang Thao Van</p>
            <p>Deputy, Sponsorship & Finance</p>
            <p>(+65) 84389906 or Hoangthaovan.pham@my.jcu.edu.au</p>
            <br/>
            <p>Akezhan Bexeitov</p>
            <p>Head, Sponsorship & Finance</p>
            <p>(+65) 82643996 or akezhan.bexeitov@my.jcu.edu.au</p>
            <br/>
            <p>for JCU Singapore 27th Convergence Conference Organizing Team</p>
            <img src="https://cdn.discordapp.com/attachments/1028583684909047808/1050069280303087686/316821900_112417125027644_1578679093283271436_n.png">
        </body>
        
        </html>
    
    """.replace('Company Name', company_name)
    message.Attachments.Add("D:\Code\Python Code\model\Email_Automation\Sponsorship Tier.pdf")
    message.Save()
    message.Send()
    time.sleep(5)




import time
from flask import Flask, render_template, request, redirect
from flask_mail import Mail, Message
from openpyxl import load_workbook, Workbook


app = Flask(__name__)
mail = Mail(app)

app.config['MAIL_SERVER'] = 'host.deca.ca'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = '<REPLACE_THIS_FROM_EMAIL_ADDRESS>'
app.config['MAIL_PASSWORD'] = '<REPLACE_THIS_EMAIL_PASSWORD>'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)
excelfile = 'directory.xlsx'
wb = load_workbook(excelfile)
sheet_obj = wb.active

@app.route("/", methods=["GET", "POST"])
def index():
    r = 2
    rows = sheet_obj.max_row
    transcript = ""
    if request.method == "POST":
        print("FORM DATA RECEIVED")
        for i in range(2, rows + 1):
            cName = sheet_obj.cell(row=r, column=1)
            pName = sheet_obj.cell(row=r, column=3)
            email = sheet_obj.cell(row=r, column=5)
            r = r + 1
            transcript = 'Dear ' + str(pName.value) + ',\n\nOur names are Agam Bhatti and Ian Korovinsky and we are Branding and Communications Officers with Ontario DECA. Established in 1979, DECA is Ontario’s premier case competition that prepares emerging leaders and entrepreneurs for careers in marketing, finance, hospitality and management in high schools and universities around the globe.\n\nOntario DECA is the largest student-run business organization in Canada. We foster the development of business knowledge and transferable skills for over 15,500 high school students across 206 schools throughout the province. Through conferences and competitions, Ontario DECA provides these truly exceptional youth with the opportunity to grow their professional network, form life-long friendships, and apply their skills to authentic business cases at an international level.\n\nWe are reaching out to you in hopes of cultivating a partnership with ' + str(cName.value) + ' over the course of the next few months. We would love to have you join the Ontario DECA community through a monetary sponsorship, where the funding would directly support our competitors through upkeep, scholarships, etc. Another viable option is an in-kind sponsorship, where you can provide your employees with the unique opportunity to inspire and direct Ontario’s emerging leaders as exhibitors and judges or provide us with resources like consumer goods and educational content that would add to the experience at our conferences.\n\nOn a more personal note, over the past few weeks, we have had the unique opportunity to learn what ' + str(cName.value) + ' is all about! We love what you have to offer and we are incredibly confident that this is something that Ontario DECA members will be interested in too!\n\nAs a partner, you will be able to connect with members, teacher advisors, parents and alumni. From exhibiting at Ontario DECA’s Provincial Competition to featured posts across our social media outlets, we provide you with a multitude of ways to interact with our network and promote what you have to offer!\n\nWe’d love to talk more with you about a partnership between ' + str(cName.value) + ' and Ontario DECA. We’ve attached a copy of our sponsorship package that contains information about our mission and purpose, member base, sponsorship tiers, and more. If you have any questions, please don’t hesitate to reach out to us.\n\nBest Regards,\nAgam Bhatti and Ian Korovinsky\nagam@deca.ca | ian@deca.ca'
            msg = Message('Ontario DECA Partnership Proposal', sender='ian@deca.ca', recipients=[email.value], cc=['agam@deca.ca'], bcc=['ian.korovinsky@outlook.com'])
            msg.body = transcript
            with app.open_resource("2022-2023 Ontario DECA Sponsorship Package.pdf") as fp:
                msg.attach("2022-2023 Ontario DECA Sponsorship Package.pdf", "application/pdf", fp.read())
            mail.send(msg)
            print("Sent")
            time.sleep(600)

    return render_template('index.html', transcript=transcript)

if __name__ == "__main__":
    app.run(debug=True, threaded=True)


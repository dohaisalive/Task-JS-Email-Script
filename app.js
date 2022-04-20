const mail = require("@sendgrid/mail");
const moment = require("moment");

const APIKEY = "secret!";

mail.setApiKey(APIKEY);

("use strict");
const excelToJson = require("convert-excel-to-json");

const result = excelToJson({
  sourceFile: "names.xlsx",
});

//console.log(result);

var today = new Date();
var dd = String(today.getDate()).padStart(2, "0");
var mm = String(today.getMonth() + 1).padStart(2, "0"); //January is 0!
var yyyy = today.getFullYear();

today = dd + "/" + mm + "/" + yyyy;

result.Sheet1.forEach((element) => {
  let firstName = element.A;
  let email = element.B;
  let courseName = element.C;
  let grade = element.D;

  let htmlCode = `
<!DOCTYPE html>
<html lang="en">
<head>
    <script src="app.js"></script>
    <script src="index.js"></script>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Certificate</title>
    <link rel="stylesheet" href="style.css">
</head>

<body>
    <div class="content" style="border-style: solid;
                                margin: auto;
                                text-align: center;
                                justify-content: center;
                                border-color: rgb(112, 9, 98);
                                padding-bottom: 2%;">
        <h1 style="color: rgb(112, 9, 98);">Certificate of Completion</h1>
        <p>This is to certify that</p>
        <h1>${firstName}</h1>
        <p>has successfully completed this course in</p>
        <h1>${courseName}</h1>
        <p>the grade is ${grade}</p>
        <div>date: ${today}</div>
        <div class="companyName">Coded</div>
    </div>
</body>
`;

  const msg = {
    to: email,
    from: {
      name: "doha almusallam",
      email: "dohaisalive@gmail.com",
    },
    subject: "certificate",

    text: "hello",
    html: htmlCode,
  };

  mail
    .send(msg)
    .then((response) => console.log("email sent"))
    .catch((error) => console.log(error.message));
});

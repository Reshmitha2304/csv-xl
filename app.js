const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const csv_parser = require('csv-parser');
// const toxls = require('json2xls');
const nodemailer = require('nodemailer');
const multer = require('multer');
var xl = require('excel4node');
const path = require('path');

var workbook = new xl.Workbook();
ws = workbook.addWorksheet("Files List");
var app = express();
var upload = multer({ dest: 'uploads/' });
var type = upload.single("file");
var content = [];
const directory = 'uploads';

var transport = nodemailer.createTransport({
    host: 'smtp.miraclesoft.com',
    port: 587,
    secure: false,
    auth: {
        user: 'rsamanthula@miraclesoft.com',
        pass: 'Mss@2k19'
    }
});

var style = workbook.createStyle({
    font: {
        color: '#ffffff',
        size: 12
    },
    fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#00aae7',
        fgColor: '#00aae7',
    }
});

app.use(function (req, res, next) {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    next();
});

app.use(bodyParser.urlencoded({
    extended: false
}));

app.use(bodyParser.json());

app.get("/",(req,res)=>{
  res.send("success");
})

app.post('/upload', type, (req, res) => {
    // console.log('inside', req.files);
    console.log('inside1', req.file);
    // console.log('inside1', req.body);
    // console.log('inside1', req);
    // res.send(req.file);
    tableData = [];
    fs.createReadStream(req.file.path)
        .pipe(csv_parser())
        .on('data', data => {
            content.push(data);
        })
        .on('end', () => {
            // // console.log('ended', content);
            // var xls = toxls(content);
            // fs.writeFileSync('data.xlsx', xls, 'binary');
            console.log(content);
            content.forEach( element =>{
                tableData.push(`<tr class="border"><td class="border">`+element.Filename+`</td>
                <td class="border">`+element.Status+`</td></tr>`)
            })
            template  = `
            <html>
              <head>
                <title>Conversion Error Report!</title>
                <link rel="stylesheet" type="text/css" href="//fonts.googleapis.com/css?family=Open+Sans">
                <style>
                  .border{
                  border-collapse:collapse;
                  border:1px solid #cdcdcd;
                  padding:5px;
                  text-align:center;
                  color:#232527;
                  font-family:Open Sans;
                  font-size:14px;
                  line-height:28px;
                  }
                  .border2{
                  border-collapse:collapse;
                  border:1px solid #cdcdcd;
                  padding:5px;
                  text-align:center;
                  color:#ffffff;
                  font-family:Open Sans;
                  font-size:14px;
                  line-height:28px;
                  font-weight:700;
                  background-color:#00aae7;
                  }
                </style>
              </head>
              <body style="background-color: #cdcdcd;margin: 0;  padding: 0; width: 100%;">
                <table>
                  <tbody>
                    <tr>
                      <td height="10px"></td>
                    </tr>
                  </tbody>
                </table>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="max-width: 800px; border-radius:13px;">
                  <tbody>
                    <tr>
                      <td>
                        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="max-width: 800px;">
                          <tbody>
                            <tr align="center">
                              <td>
                                <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#232527" style=" border-radius:7px 7px 0 0;background-color:#232527;max-width: 800px;background-size: cover;background-repeat: no-repeat;">
                                  <tbody>
                                    <tr>
                                      <td height="5px"></td>
                                    </tr>
                                    <tr>
                                      <td align="center">
                                        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="max-width: 700px;">
                                          <tbody>
                                            <tr>
                                              <td>
                                                <table width="100%" border="0" cellpadding="0" cellspacing="0" style="max-width: 700px;">
                                                  <tbody>
                                                    <tr>
                                                      <td valign="middle" height="auto" width="160" style="padding: 0 7px; text-align:center">
                                                        <a href="http://me.miraclesoft.com/" target="_blank">
                                                        <img style="width: 70px;" src="http://www.miraclesoft.com/images/newsletters/2017/October/MiracleMe_logo.png"></a>
                                                      </td>
                                                      <td width="100%" style="font-family:Open Sans; color:#000000;" align="Right">
                                                        <span style="font-size:14px;font-family:Open Sans;text-decoration: none;color:#ffffff;">
                                                        <a href="http://www.miraclesoft.com/company/" target="_blank" style="text-decoration: none; color:#ffffff;font-weight: 500;">About</a>
                                                        </span>
                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                        <span style="font-family:Open Sans;font-size:14px;text-decoration: none; color:#ffffff;">
                                                        <a href="http://www.miraclesoft.com/services/" target="_blank" style="text-decoration: none;color:#ffffff;font-weight: 500;"> Services</a>
                                                        </span>
                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                        <span style="font-family:Open Sans;font-size:14px;text-decoration: none; color:#ffffff;">
                                                        <a href="http://www.miraclesoft.com/contact/" target="_blank" style="text-decoration: none;color:#ffffff;font-weight: 500;">Contact
                                                        </a>
                                                        </span>
                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                      </td>
                                                    </tr>
                                                  </tbody>
                                                </table>
                                              </td>
                                            </tr>
                                          </tbody>
                                        </table>
                                      </td>
                                    </tr>
                                    <tr>
                                      <td height="5px"></td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="max-width:800px;background-position: center top; background-repeat: no-repeat; background-size: cover; background-color: #ffffff;" bgcolor="#ffffff">
                          <tbody>
                            <tr align="center">
                              <td>
                                <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#232527" style="background-color:#f1f1f1;max-width: 800px;background-size: cover;background-repeat: no-repeat;" background="http://d2b8lqy494c4mo.cloudfront.net/mss/images/newsletters/2018/May/mtalk_banner1_may18.png">
                                  <tbody>
                                    <tr>
                                      <td height="15px"></td>
                                    </tr>
                                    <tr>
                                      <td align="center">
                                        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="max-width: 700px;/* padding: 10px; */">
                                          <tbody>
                                            <tr>
                                              <td>
                                                <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="">
                                                  <tbody>
                                                    <tr>
                                                      <td style="">
                                                        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
                                                          <tbody>
                                                            <tr>
                                                              <td height="5px"></td>
                                                            </tr>
                                                            <tr>
                                                              <td style="font-family:'Open Sans';font-size: 35px;line-height: 50px;font-weight: 800;color: #ffffff;" align="left">
                                                                <b>PDF Conversion Error Report!</b>
                                                              </td>
                                                            </tr>
                                                            <tr>
                                                              <td height="5px"></td>
                                                            </tr>
                                                          </tbody>
                                                        </table>
                                                      </td>
                                                    </tr>
                                                  </tbody>
                                                </table>
                                              </td>
                                            </tr>
                                          </tbody>
                                        </table>
                                      </td>
                                    </tr>
                                    <tr>
                                      <td height="15px"></td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                            <tr>
                              <td align="center">
                                <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="
                                  max-width: 700px;
                                  ">
                                  <tbody>
                                    <tr>
                                      <td>
                                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                          <tbody>
                                            <tr>
                                              <td style="">
                                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                  <tbody>
                                                    <tr>
                                                      <td height="20px"></td>
                                                    </tr>
                                                    <tr>
                                                      <td style="padding: 0 5px;text-align:left;font-family: 'Open Sans';font-size:15px;line-height: 25px;text-decoration: none;color: #232527;font-weight:400;">
                                                        <b>Hello Team,</b>
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <td height="15px"></td>
                                                    </tr>
                                                    <tr>
                                                      <td style="padding: 0 5px;text-align:justify; font-family: 'Open Sans'; font-size:15px; line-height: 25px; text-decoration: none; color: #232527; font-weight:500;">
                                                        The PPT to PDF conversion of the requested documents has been completed. Please check the below generated report to know the conversion status of all the documents.
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <td height="15px"></td>
                                                    </tr>
                                                    <tr>
                                                      <td style="padding:0 5px;color:#232527;font-family:Open Sans;font-size:14px;line-height:28px" align="justify">
                                                        <table width="100%" class="border">
                                                          <tr class="border2">
                                                            <th class="border2"><b>Document Title</b></th>
                                                            <th class="border2"><b>Status Report</b></th>
                                                            </tr class="border">
                                                            `+tableData+`
                                                        </table>
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <td height="20px"></td>
                                                    </tr>
                                                  </tbody>
                                                </table>
                                              </td>
                                            </tr>
                                          </tbody>
                                        </table>
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" border="0" bgcolor="#232527" align="center" width="100%" style="max-width: 800px; border-radius:0 0 7px 7px;">
                          <tbody>
                            <tr>
                              <td align="center">
                                <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%" style="
                                  max-width: 700px;
                                  ">
                                  <tbody>
                                    <tr>
                                      <td height="10px"></td>
                                    </tr>
                                    <tr>
                                      <td align="center">
                                        <table style="max-width: 360px; display: inline-block;" cellspacing="0" cellpadding="0" border="0" align="left" width="100%">
                                          <tbody>
                                            <tr>
                                              <td height="5"></td>
                                            </tr>
                                            <tr>
                                              <td style="color: rgb(255, 255, 255); font-family: 'open sans'; font-size: 13px; line-height: 23px; font-weight: 500; padding: 8px;" align="center">
                                                &#169; Copyrights 2019 | Miracle Software Systems, Inc.
                                              </td>
                                            </tr>
                                          </tbody>
                                        </table>
                                        <table style="max-width: 210px; display: inline-block; padding: 10px;" cellspacing="0" cellpadding="0" border="0" align="right" width="100%">
                                          <tbody>
                                            <tr>
                                              <td style="font-size:0!important" align="center" width="100%">
                                                <div style="width:100%;display:inline-block;vertical-align:top;font-size:normal;max-width:50px">
                                                  <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                    <tbody>
                                                      <tr>
                                                        <td style="padding:0 10px" align="center">
                                                          <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                            <tbody>
                                                              <tr>
                                                                <td style="line-height:0!important" align="center">
                                                                  <a style="text-decoration:none;display:inline-block" href="https://facebook.com/miracle45625" target="_blank">
                                                                  <img alt="socials1" width="30" src="http://d2b8lqy494c4mo.cloudfront.net/mss/images/newsletters/2018/February/facebook-logo-button.png">
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
                                                        </td>
                                                      </tr>
                                                    </tbody>
                                                  </table>
                                                </div>
                                                <div style="width:100%;display:inline-block;vertical-align:top;font-size:normal;max-width:50px">
                                                  <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                    <tbody>
                                                      <tr>
                                                        <td style="padding:0 10px" align="center">
                                                          <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                            <tbody>
                                                              <tr>
                                                                <td style="line-height:0!important" align="center">
                                                                  <a style="text-decoration:none;display:inline-block" href="https://www.instagram.com/team_mss/" target="_blank">
                                                                  <img alt="socials1" width="30" src="https://d2b8lqy494c4mo.cloudfront.net/mss/images/newsletters/2019/April/instagram_new_white.png">
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
                                                        </td>
                                                      </tr>
                                                    </tbody>
                                                  </table>
                                                </div>
                                                <div style="width:100%;display:inline-block;vertical-align:top;font-size:normal;max-width:50px">
                                                  <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                    <tbody>
                                                      <tr>
                                                        <td style="padding:0 10px" align="center">
                                                          <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                            <tbody>
                                                              <tr>
                                                                <td style="line-height:0!important" align="center">
                                                                  <a style="text-decoration:none;display:inline-block" href="https://www.linkedin.com/company/miracle-software-systems-inc" target="_blank">
                                                                  <img alt="socials1" width="30" src="http://d2b8lqy494c4mo.cloudfront.net/mss/images/newsletters/2018/February/linkedin-button.png">
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
                                                        </td>
                                                      </tr>
                                                    </tbody>
                                                  </table>
                                                </div>
                                                <div style="width:100%;display:inline-block;vertical-align:top;font-size:normal;max-width:50px">
                                                  <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                    <tbody>
                                                      <tr>
                                                        <td style="padding:0 10px" align="center">
                                                          <table cellspacing="0" cellpadding="0" border="0" align="center" width="100%">
                                                            <tbody>
                                                              <tr>
                                                                <td style="line-height:0!important" align="center">
                                                                  <a style="text-decoration:none;display:inline-block" href="https://www.youtube.com/c/Team_MSS" target="_blank">
                                                                  <img alt="socials3" width="30" src="http://d2b8lqy494c4mo.cloudfront.net/mss/images/newsletters/2018/February/youtube-logotype.png">
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
                                                        </td>
                                                      </tr>
                                                    </tbody>
                                                  </table>
                                                </div>
                                              </td>
                                            </tr>
                                          </tbody>
                                        </table>
                                      </td>
                                    </tr>
                                    <tr>
                                      <td height="10px"></td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <table>
                  <tbody>
                    <tr>
                      <td height="10px"></td>
                    </tr>
                  </tbody>
                </table>
              </body>
            </html>
            `
            var count = 0;
            ws.cell(1, 1).string("S.No.").style(style);
            ws.cell(1, 2).string("File Name").style(style);
            ws.cell(1, 3).string("Status").style(style);

            for (i = 0; i < content.length; i++) {
                count++;
                ws.cell(count + 1, 1).number(count);
                ws.cell(count + 1, 2).string(content[i].Filename);
                ws.cell(count + 1, 3).string(content[i].Status);
            }
        //    officersIds = content.map(officer => officer);
        //     console.log(officersIds);
            workbook.writeToBuffer().then((buffer) => {
                //  workbook.write('data.xlsx', buffer);
                let mailOptions = {
                    from: 'rsamanthula@miraclesoft.com', // sender address
                    to: ['rsamanthula@miraclesoft.com'],
                    subject: 'PDF Conversion Error Report!', // Subject line
                    html: template // html body,
                    // attachments: [
                    //     {
                    //         filename: 'data123.xlsx',
                    //         // path: __dirname + '/data.xlsx'
                    //         content: buffer
                    //     }
                    // ]
                };

                transport.sendMail(mailOptions, (error, info) => {
                    if (!error) {
                        content = [];
                        count = 0;
                        fs.readdir(directory, (err, files) => {
                            if (err) throw err;

                            for (const file of files) {
                                fs.unlink(path.join(directory, file), err => {
                                    if (err) throw err;
                                    res.send({ 'status': 'success' });
                                });
                            }
                        });
                    } else {
                        res.send(error);
                        console.log(error)
                    }
                });
            });
        });
});

// app.listen(3000, () => {
//     console.log("on 3000");
// });

http.listen(process.env.PORT || 3000, function(){
  console.log('listening on', http.address().port);
});

// try {
//     // 18 minutes ago (from now)
//            var query =   {
//                $where: "this.fileVersions[this.fileVersions.length - 1].fileType !== 'pdf'"
//              }
//            console.log("inside of funcation",query);
//            seperateQueryDb(query).then(data => {
//               console.log("status of response", data);
//                callback(data)

//              });

//          } catch (err) {
//            next(err)
//          }

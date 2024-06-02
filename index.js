const multer = require("multer");
const xlsx = require('xlsx');
const express = require('express');
const fs = require('fs');
const archiver = require('archiver');
const htmlDocx = require('html-docx-js');



const app = express();
const port  = 3000;
let fileName;
let data;

const ds = multer.diskStorage({
    destination:'import/',
    filename:(req,file,cb) => {
        fileName = file.originalname;
        cb(null,file.originalname);
    }
});

const a = multer({
    storage:ds
});

function convertIntoJSON(){
    const workbook = xlsx.readFile(`import/${fileName}`);

    // Assume data is in the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert worksheet to JSON
    const jsonData = xlsx.utils.sheet_to_json(worksheet);

    // Write JSON data

    return JSON.stringify(jsonData);
}
app.use(express.json());
app.use(express.static("public"));

app.get("/",(req,res) => {
    res.render("index.ejs");
});

app.post("/upload",a.single('excelFile'),(req,res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    } 
     
      let jsons = convertIntoJSON();
      data = JSON.parse(jsons);
      const { Name,Clzname,rno,title,startdate,enddate,branch,gender } = data[0];
      if(!branch || !Name || !Clzname || !rno || !title || !startdate || !enddate || !gender){
        const reasonObj = {b:branch,n:Name,c:Clzname,r:rno,t:title,s:startdate,e:enddate,g:gender};
        res.render("error.ejs",reasonObj);
      }else{
      fs.readFile('public/success.html', 'utf8', (err, data) => {
        if (err) {
          return res.status(500).send('Error reading HTML file.');
        }
        res.send(data);
      });
    }
});

app.get('/download', async(req, res) => {
    
    var htmlFiles = [];
    
    function convertIntoDate(dateToBeConverted){
         // Serial number from Excel
    var serialNumber = dateToBeConverted;

    // Number of milliseconds in a day
    var millisecondsPerDay = 24 * 60 * 60 * 1000;

    // Base date in milliseconds in Excel (December 30, 1899)
    var baseDateMilliseconds = new Date(1899, 11, 30).getTime();

    // Calculate the milliseconds offset for the serial number
    var serialNumberOffsetMilliseconds = serialNumber * millisecondsPerDay;

    // Calculate the resulting date in milliseconds
    var resultingDateMilliseconds = baseDateMilliseconds + serialNumberOffsetMilliseconds;

    // Create a new Date object with the resulting date
    var resultingDate = new Date(resultingDateMilliseconds);
    var cdy = resultingDate.getFullYear();
    var cdm = resultingDate.getMonth() + 1;
    var cdd = resultingDate.getDate();
    if(cdm < 10){
        cdm = "0" + cdm;
    }
    if(cdd < 10){
        cdd = "0" + cdd;
    }
    var convertedDate = cdd+"-"+cdm+"-"+cdy;
    return convertedDate;
    }

    data.forEach((element, index) => {

        let { Name,Clzname,rno,title,startdate,enddate,branch,gender } = element;
        Name = Name.trim().toUpperCase();
        Clzname = Clzname.trim().toUpperCase();
        title = title.trim();
        branch = branch.trim();
        var genderLabel1;
        var genderLabel2;

        if(gender.trim().toLowerCase() == 'm'){
            genderLabel1 = "Mr";
            genderLabel2 = "He";
        }else if(gender.trim().toLowerCase() == 'f'){
            genderLabel1 = "Ms";
            genderLabel2 = "She";
        }else{
            genderLabel1 = "Mr/Ms";
            genderLabel2 = "He/She";
        }

        var d = new Date();
        var y = d.getFullYear();
        var dad = d.getDate();
        var m = d.getMonth() + 1;
        if(m < 10){
            m = "0"+m;
        }
        if(dad < 10){
            dad = "0"+dad;
        }
        var todayDate = dad+"-"+m+"-"+y; 

        const htmlTemplate = `<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                
                .dtp {
                    text-align: right;
                    margin: 0;
                    margin-top: 20px; 
                }
                #dt{
                    margin-bottom: 10px;
                    font-family: 'Times New Roman', Times, serif;
                    font-size:20px;
                }
                
                h2{
                    text-align: center;
                    margin-top: 60px;
                    font-family: "Calibri",sans-serif;
                    font-size: 30px;
                    font-weight: bold;
                }
                #sign {
                    margin-top: 50px;
                    margin-bottom: 10px;
                }
                #sub {
                    font-size: 1.5em;
                    margin-left: 50px;
                    margin-top: 60px;
                    font-family: 'Times New Roman', Times, serif;
                }
                 #matter {
                    margin-top: 50px;
                    font-family: "Calibri",sans-serif;
                    font-size: 20px;
                    text-align: justify;
                }
               
                p {
                    margin: 0 auto;
                    font-size: 20px;
                    max-width: 800px;
                    margin-bottom: 29px;
                    font-style: normal;
                    line-height: 170%;
                }
                .comp{
                    text-align: right;
                    margin-right: 20px;
                    margin-left: 5rem;
                   
                }
                .sign{
                    font-size: 20px;
                    margin-top: 70px;
                    font-family: 'Times New Roman', Times, serif;
                }
                 .this{
                    margin-left:50px;
                 }  
                 .first-line{
                    text-indent:50px;
                 }  
                 .space{
                    margin-top:120px;
                 }        
            </style>
        </head>
        <body>
            <div class="space"> 
            <div>
            <div class="dtp">
                <p id="dt">Date: ${todayDate}</p>
                
            </div>
            <h2>CERTIFICATE OF INTERNSHIP</h2>
            <p id="sub">Sub: Successful completion of Internship</p>
            

            <p id="matter" class="first-line">This is to certify that ${genderLabel1} <b>${Name}</b> a student of ,<b>${Clzname}</b>
            bearing the <b>Reg. No ${rno}</b> has successfully completed project in
            <b>${title}</b> at <b>SYMBIOSYS TECHNOLOGIES</b> from ${convertIntoDate(startdate)} to ${convertIntoDate(enddate)}
            under our guidance. ${genderLabel2} has exhibited very good analytical skills and demonstrated
            good technical understanding</p>
            </div>
            <table class="sign">
            <tr>
                <td>Project Manager</td>
                <td style="padding-left: 300px;">HR Manager</td>
            </tr>
        </table>
        </div>
        </body>
        </html>`;
        htmlFiles.push(htmlTemplate);
    });

    try {
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; ');

        const archive = archiver('zip', {
            zlib: { level: 9 }
        });

        archive.on('error', function(err) {
            throw err;
        });

        // Pipe the archive data to the response
        archive.pipe(res);

        for (let index = 0; index < htmlFiles.length; index++) {
            const htmlContent = htmlFiles[index];
            const docxContent = htmlDocx.asBlob(htmlContent);
            const buffer = Buffer.from(await docxContent.arrayBuffer()); // Convert Blob to Buffer
            archive.append(buffer, { name: `document${index + 1}.docx` });
        }

        await archive.finalize();
    } catch (error) {
        console.error('Error generating documents:', error);
        res.status(500).send('Internal Server Error');
    }


});

app.get("/offer",async (req,res) => {

    var d = new Date();
    var y = d.getFullYear();
    var dad = d.getDate();
    var m = d.getMonth() + 1;
    if(m < 10){
        m = "0"+m;
    }
    if(dad < 10){
        dad = "0"+dad;
    }
    let todayDate = dad+"-"+m+"-"+y;
    
    let obj = {};
    data.forEach((ele) => {
        let {Clzname,branch} = ele;

        Clzname = Clzname.trim().toLowerCase();
        branch = branch.trim().toLowerCase();
       
        if(obj[Clzname]){
            if(obj[Clzname][branch]){
                obj[Clzname][branch].push(ele);
            }else{
                let arr = [];
                arr.push(ele);
                Object.assign(obj[Clzname],{[branch]:arr});
            }
        }else{
            let arr = [];
            arr.push(ele);
            Object.assign(obj, { [Clzname]: {[branch] : arr} });
        }
    });

    function ltten(studentBranch,studentCollegeName,s,todayDate){
        let htmlTemplate2 =`<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                
                .dtp {
                    text-align: right;
                    margin: 0;
                    margin-top: 20px; 
                }
                #dt{
                    margin-bottom: 10px;
                }
               
                h3 {
                    text-align: center;
                    margin-top: 50px;
                }
                #sign {
                    margin-bottom: 10px;
                }
                #sub {
                    margin-top: 50px;
                }
                 #matter {
                    margin-top: 25px;
                }
                table {
                    background-color: #ffffff;
                    border-collapse: collapse;
                    border-width: 2px;
                    border-color: #000000;
                    border-style: solid;
                    color: #000000;
                    }

                    table td, table th {
                    border-width: 2px;
                    border-color: #000000;
                    border-style: solid;
                    padding: 3px;
                    }

                    table thead {
                    background-color: #ffffff;
                }
                table th,table td,table tr{
                    padding: 1px 40px;
                }
                p {
                    margin: 0 auto;
                    max-width: 800px;
                    margin-bottom: 29px;
                }
                #tnq{
                    margin-top: 40px;
                    margin-bottom: 10px;
                } 
                
                p{
                    font-style: normal;
                }
            </style>
        </head>
        <body>
            <div class="dtp">
                <p id="dt">Date: ${todayDate}</p>
                <p id="plc">Visakhapatnam,</p>
            </div>
            <h3>TO WHOM SO EVER IT MAY CONCERN</h3>
            <p id="sub">Sub: Internship Acceptance Letter,</p>
            <p id="matter">This is reference to your letter dated on ${todayDate}. We are pleased to offer you an internship program with <b>SYMBIOSYS TECHNOLOGIES</b> for 45 days for below ${studentBranch} student from ${studentCollegeName}</p>
            <table border="1" align="center" >
                <tr><th>Sno</th><th>Name</th><th>RegNo</th><th>Branch</th></tr>
                ${s}
            </table>
            <div>
            <p id="tnq">Thanking You,</p>
            <p style="margin-bottom: 40px;">Yours Sincerely,</p>
            <p id="sign">D.Sudheer</p>
            <p>Authorized Signatory</p>
            </div>    
        </body>
        </html>`;
            s="";
            htmlFiles.push(htmlTemplate2);
    }

    var s="";
    var htmlFiles = [];
    for(var key in obj) {
        var value = obj[key];
        for(var k in value){
            var v = obj[key][k];
            var len = v.length;
            var studentBranch;
            var studentCollegeName;
            var c=0;
            for(var i=0;i<len;i++){
                c++;
                studentBranch = v[i]['branch'];
                studentCollegeName = v[i]['Clzname'];
                s = s + `<tr><td>${i+1}</td><td>${v[i]['Name']}</td><td>${v[i]['rno']}</td><td>${v[i]['branch']}</td></tr>`;
                if(c==8||i==len-1){
                    ltten(studentBranch,studentCollegeName,s,todayDate);
                    s="";
                    c=0;
                }
            }
           
        }
    }

    try {
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; ');

        const archive = archiver('zip', {
            zlib: { level: 9 }
        });

        archive.on('error', function(err) {
            throw err;
        });

        // Pipe the archive data to the response
        archive.pipe(res);

        for (let index = 0; index < htmlFiles.length; index++) {
            const htmlContent = htmlFiles[index];
            const docxContent = htmlDocx.asBlob(htmlContent);
            const buffer = Buffer.from(await docxContent.arrayBuffer()); // Convert Blob to Buffer
            archive.append(buffer, { name: `offerLetters${index + 1}.docx` });
        }

        await archive.finalize();
    } catch (error) {
        console.error('Error generating documents:', error);
        res.status(500).send('Internal Server Error');
    }

});

app.get("/contact",(req,res) => {
    res.render("contact.ejs");
});

app.listen(port,() => {
 console.log(`Server running on port ${port}`);
 console.log(" ");
 console.log(`Click here(Ctrl + Click): http://localhost:3000/`);
 console.log(" ");
 console.log(`To stop the execution press: "Ctrl + C" and then click 'Y' or 'y'`);
});
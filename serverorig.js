const express = require("express");
const excel = require("exceljs");
var mysql = require('mysql');
const body = require("body-parser");
var connection = mysql.createConnection({
    host:'localhost',
    user:'root',
    password:'',
    database:'nptel'
})

connection.connect(function(error){

    if(!!error){
        console.log('Error'); 
       } else{

        console.log('Connected');
       }
});
var app = express();
app.use(express.static("public"));
app.use(body.urlencoded({ extended: false }))

app.post("/formlogin", function (req, res) {
    var name = req.body.first_name;

    var dbexcel = {
        name: req.body.first_name,
        lastname: req.body.last_name,
        birthday: req.body.birthday,
        roll: req.body.roll,
        caste: req.body.caste,
        gender: req.body.gender,
        email: req.body.email,
        phone: req.body.phone,
		year:'2020',
        Class: req.body.year,
        branch: req.body.branch,
        position: req.body.position,
        coursename: req.body.coursename

    }

    console.log(dbexcel.name,dbexcel.lastname,dbexcel.birthday,dbexcel.roll,dbexcel.caste,dbexcel.gender,dbexcel.email,dbexcel.phone,dbexcel.year,dbexcel.branch,dbexcel.position,dbexcel.coursename);
console.log(dbexcel.name);
console.log(typeof(dbexcel.roll));
   // var sql = "INSERT INTO students VALUES ("+dbexcel.name+","+ dbexcel.lastname+","+dbexcel.birthday+","+dbexcel.roll+","+dbexcel.caste+","+dbexcel.gender+","+dbexcel.email+","+dbexcel.phone+","+dbexcel.year+","+dbexcel.branch+","+dbexcel.position+","+dbexcel.coursename+")";
//var sql ="INSERT INTO students VALUES('aditya','Jain','23-04-2009','18','no','M','email@gmail.com','872392931','TE','IT','Student','Data Binding')";
   var sql ="INSERT INTO registration VALUES ?";
   var records=[[dbexcel.name,dbexcel.lastname,dbexcel.birthday,dbexcel.roll,dbexcel.caste,dbexcel.gender,dbexcel.email,dbexcel.phone,dbexcel.year,dbexcel.branch,dbexcel.position,dbexcel.coursename,dbexcel.Class]];
     
   connection.query(sql,[records],function(error,rows,fields){
        if(!!error){
            console.log(error); 
           } else{
            console.log('success');
           }
   
        });

    res.redirect("/loginsuccess");
});

app.get("/loginsuccess", function (req, res) {
    let conn;
    var data1;
    
    connection.query("SELECT * FROM registration", function (err, students, fields) {
    
        const jsonStudents = JSON.parse(JSON.stringify(students));
        console.log(jsonStudents);
   
    console.log("I am in Loginsuccess function")
    let workbook = new excel.Workbook();
    let worksheet = workbook.addWorksheet('Customers'); //creating worksheet
   
    worksheet.columns = [
        { header: 'Name', key: 'Name', width: 10 },
        { header: 'Lastname', key: 'Lastname', width: 10 },
        { header: 'DOB', key: 'DOB', width: 10 },
        { header: 'Rollno', key: 'Rollno', width: 10 },
        { header: 'Caste', key: 'Caste', width: 10 },
        { header: 'Gender', key: 'Gender', width: 10 },
        { header: 'Email', key: 'Email', width: 10 },
        { header: 'Phone', key: 'Phone', width: 10 },
        { header: 'Year', key: 'Year', width: 10 },
        { header: "Branch", key: 'Branch', width: 20 },
        { header: 'Student/Faculty', key: 'Student/Faculty', width: 10 },
        { header: 'Coursename', key: 'Coursename', width: 10 },
    ];

    worksheet.addRows(jsonStudents);
    workbook.xlsx.writeFile("customer.xlsx")
        .then(function () {
            console.log("file saved!");
        });
    });
    res.send("<h1>Done</h1>");

});

app.listen(8080, function () {
    console.log("Server is running on port 8080");
});
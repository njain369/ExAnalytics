const express = require("express");
const excel = require("exceljs");
const body = require("body-parser");
var mongo = require("mongodb").MongoClient;
var url = "mongodb://localhost:27017";

var app = express();
app.use(express.static("public"));
app.use(body.urlencoded({ extended: false }))
function init(){    
    let conn;
    mongo.connect(url)
            .then((client) => {
              conn = client;
              return client.db("ExAnalytics");
            })
            .then((db) => db.collection("Students"))
            .then((collection)=>collection.find())
            .then(cursor=>cursor.toArray())
            .then(data=>app.locals.users=data)
            .then(() => conn.close());
}
init();
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
        year: req.body.year,
        branch: req.body.branch,
        position: req.body.position,
        coursename: req.body.coursename

    }


    let conn;
   (()=>{ mongo.connect(url)
        .then((client) => {
            conn = client;
            return client.db("ExAnalytics");
        })
        .then((db) => db.collection("Students"))
        .then((collection) => { collection.insertOne(dbexcel) })
        .then(() => conn.close())
    })();
   
    function init(){    
        let conn;
        mongo.connect(url)
                .then((client) => {
                  conn = client;
                  return client.db("ExAnalytics");
                })
                .then((db) => db.collection("Students"))
                .then((collection)=>collection.find())
                .then(cursor=>cursor.toArray())
                .then(data=>app.locals.users=data)
                .then(() => conn.close());
    }
    init();

    res.redirect("/loginsuccess");
});

app.get("/loginsuccess", function (req, res) {
    let conn;
    var data1;
    mongo.connect(url)
        .then((client)=>{
            conn=client;
            return client.db("ExAnalytics");
        })
        .then((db)=>db.collection("Students"))
        .then(collection=>collection.find())
        .then(cursor=>cursor.toArray())
        .then((data)=>{
            app.locals.users=data;
        })
        .then(()=>conn.close())
    

        
    console.log(app.locals.users);
var data1=app.locals.users;
    
    console.log("I am in Loginsuccess function")
    let workbook = new excel.Workbook();
    let worksheet = workbook.addWorksheet('Customers'); //creating worksheet
   
    worksheet.columns = [
        { header: 'Name', key: 'name', width: 10 },
        { header: 'Lastname', key: 'lastname', width: 10 },
        { header: 'DOB', key: 'birthday', width: 10 },
        { header: 'Rollno', key: 'roll', width: 10 },
        { header: 'Caste', key: 'caste', width: 10 },
        { header: 'Gender', key: 'gender', width: 10 },
        { header: 'Email', key: 'email', width: 10 },
        { header: 'Phone', key: 'phone', width: 10 },
        { header: 'Year', key: 'year', width: 10 },
        { header: "Branch", key: 'branch', width: 20 },
        { header: 'Student/Faculty', key: 'position', width: 10 },
        { header: 'Coursename', key: 'coursename', width: 10 },
    ];

    worksheet.addRows(data1);
    workbook.xlsx.writeFile("customer.xlsx")
        .then(function () {
            console.log("file saved!");
        });
    res.send("<h1>Done</h1>");

});

app.listen(8080, function () {
    console.log("Server is running on port 8080");
});
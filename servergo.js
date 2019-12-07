var express=require("express");
var app=express();
var bodyParser=require("body-parser");

app.use(bodyParser.urlencoded({extended:false}));

app.use(express.static("public"));
app.post("/one",function(req,res){
    var value=req.body.value;
    console.log(value);
})

app.listen(8080,function(){
    console.log("Server is on at port 8080");
})
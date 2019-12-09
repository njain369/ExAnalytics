import React, { Component } from 'react';
import './App.css';
import {OutTable, ExcelRenderer} from 'react-excel-renderer';
import { Jumbotron, Col, Input, InputGroup, InputGroupAddon, FormGroup, Label, Button, Fade, FormFeedback, Container, Card } from 'reactstrap';

class App extends Component {

  constructor(props){
    super(props);
    this.state={
      isOpen: false,
      dataLoaded: false,
      isFormInvalid: false,
      rows: null,
      cols: null
    }
    this.fileHandler = this.fileHandler.bind(this);
    this.toggle = this.toggle.bind(this);
    this.openFileBrowser = this.openFileBrowser.bind(this);
    this.sqlupload = this.sqlupload.bind(this);
    this.renderFile = this.renderFile.bind(this);
    this.openNewPage = this.openNewPage.bind(this);
    this.fileInput = React.createRef();
  }

  sqlupload = (event) => {

    this.setState({
      rows : event.target.value
    });
    
    console.log("Done");
    console.log(this.state.rows);

//     var express = require('express');
//   var mysql = require('mysql');
//   var app = express();
//   var connection = mysql.createConnection({
//     host:'localhost',
//     user:'root',
//     password:'',
//     database:'sampledb'
// })
//   connection.connect(function(error){

//       if(!!error){
//           console.log('Error'); 
//         } else{

//           console.log('Connectd');
//         }
//   });
//   app.get('/',function(req,resp){
//     var sql = "INSERT INTO mysampletable VALUES ('aditya', 'khati', 'amkhati')";
//   connection.query(sql,function(error,rows,fields){
//       if(!!error){
//           console.log('Error in the query'); 
//         } else{
//           console.log('sucess');
//         }
//   });
//   });
  }

  renderFile = (fileObj) => {
      //just pass the fileObj as parameter
      ExcelRenderer(fileObj, (err, resp) => {
        if(err){
          console.log(err);            
        }
        else{
          this.setState({
            dataLoaded: true,
            cols: resp.cols,
            rows: resp.rows
          });
        }
      }); 
  }

  fileHandler = (event) => {    
    if(event.target.files.length){
      let fileObj = event.target.files[0];
      let fileName = fileObj.name;

      
      //check for file extension and pass only if it is .xlsx and display error message otherwise
      if(fileName.slice(fileName.lastIndexOf('.')+1) === "xlsx"){
        this.setState({
          uploadedFileName: fileName,
          isFormInvalid: false
        });
        this.renderFile(fileObj)
      }    
      else{
        this.setState({
          isFormInvalid: true,
          uploadedFileName: ""
        })
      }
    }               
  }

  toggle() {
    this.setState({
      isOpen: !this.state.isOpen
    });
  }

  openFileBrowser = () => {
    this.fileInput.current.click();
  }

  openNewPage = (chosenItem) => {
    const url = chosenItem === "github" ? "https://github.com/ashishd751/react-excel-renderer" : "https://medium.com/@ashishd751/render-and-display-excel-sheets-on-webpage-using-react-js-af785a5db6a7";
    window.open(url, '_blank');
  }

  render() {
    return (
      <div style={{marginTop:"100px"}}>
        <Container>
        <form>
          <FormGroup row>
            <Label for="exampleFile" xs={6} sm={4} lg={2} size="lg">Upload</Label>          
            <Col xs={4} sm={8} lg={10}>                                                     
              <InputGroup>
                <InputGroupAddon addonType="prepend">
                  <Button color="info" style={{color: "white", zIndex: 0}} onClick={this.openFileBrowser.bind(this)}><i className="cui-file"></i> Browse&hellip;</Button>
                  <input type="file" hidden onChange={this.fileHandler.bind(this)} ref={this.fileInput} onClick={(event)=> { event.target.value = null }} style={{"padding":"10px"}} />                                
                </InputGroupAddon>
                <Input type="text" className="form-control" value={this.state.uploadedFileName} readOnly invalid={this.state.isFormInvalid} />                                              
                <FormFeedback>    
                  <Fade in={this.state.isFormInvalid} tag="h6" style={{fontStyle: "italic"}}>
                    Please select a .xlsx file only !
                  </Fade>                                                                
                </FormFeedback>
              </InputGroup>     
            </Col>                                                   
          </FormGroup>   
        </form>

        {this.state.dataLoaded && 
        <div>
          <Card body outline color="secondary" className="restrict-card">
            
              <OutTable data={this.state.rows} columns={this.state.cols} tableClassName="ExcelTable2007" tableHeaderRowClass="heading" />
            
          </Card>  
        </div>}
        <Button color="info" style={{color: "white", zIndex: 0, margin:"auto",marginTop:"30px", display:"block"}} onClick={this.sqlupload.bind(this)} value={this.state.renderFile}><i className="cui-file"></i> Upload</Button>

        </Container>
      </div>
    );
  }
}

export default App;

import express from 'express';
import mongoose , { Model } from 'mongoose';
import bodyParser from 'body-parser';
import path, { dirname } from "path";
import { fileURLToPath } from "url";
import XLSX from 'xlsx';
import multer from 'multer';
import stringify from 'querystring';

//multer
const __fileName = fileURLToPath(import.meta.url);
const __dirname = dirname(__fileName);
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './public/uploads')
    },
    filename: function (req, file, cb) {
      const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9)
      cb(null, file.originalname)
    }
  });
  
  const upload = multer({ storage: storage });

//connecting to db
mongoose.connect('mongodb+srv://Saksham2002:saksham@cluster0.dbqjgv3.mongodb.net')
.then(()=>{console.log('connected to db')})
.catch((error)=>{console.log('error',error)});

//init app
const app = express();

//setting the template engine
app.set('view engine','ejs');

//fetching data from the request
app.use(bodyParser.urlencoded({extended:false}));

//static folder path
app.use(express.static(path.resolve(__dirname,'public')));

//collection schema
const excelSchema = new mongoose.Schema({
    ["Name of the Candidate"]:String,
    ["Email"]:String,
    ["Mobile No"]:String,
    ["Date of Birth"]:String,
    ["Work Experience"]:String,
    ["Resume Title"]:String,
    ["Current Location"]:String,
    ["Postal Address"]:String,
    ["Current Employer"]:String,
    ["Current Designation"]:String,
});

const excelModel = mongoose.model('excelData',excelSchema);

app.get('/',(req,res)=>{
    excelModel.find((err,data)=>{
        if(err){
            console.log(err)
        }else{
            if(data!=''){
                //console.log(data);
                res.render('home',{result:data});
            }else{
                res.render('home',{result:{}});
            }
        }
    });
});

app.post('/',upload.single('excel'),(req,res)=>{
    const workbook = XLSX.readFile(req.file.path);
    const sheet_namelist = workbook.SheetNames;
    console.log(sheet_namelist);
    sheet_namelist.forEach((element,x) => {
        let xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[x]],{raw: false});
        xlData = xlData.map(obj => ({...obj, 'Mobile No': obj['Mobile No.']}))
        excelModel.insertMany(xlData,(err,data)=>{
            if(err){
                console.log(err);
            }else{
               console.log(data);
            }
        })

        //console.log(xlData[0]["Mobile No."])
        //console.log(xlData);
    });
    res.redirect('/');
});

//assign port
const port = process.env.PORT || 3000;
app.listen(port,()=>console.log('server run at'+port)); 
const { default: axios } = require('axios');
const express=require('express');
const multer = require('multer');
const xlsx= require('xlsx');
const path=require('path')

var app = express();

app.use(express.json());

//Multer Config for storing the sheet 
const fileStorageEngine = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, "./sheet"); 
    },
    filename: (req, file, cb) => {
      cb(null,file.originalname);
      fileName=file.originalname;
    },
  });


const upload=multer({storage : fileStorageEngine})

// Storing for product details
const product_details = new Map();

app.post("/sendFile", upload.single("sheet"), async (req, res) => {

   // Getting the JSON Data from the sheet
   const wb= xlsx.readFile('./sheet/product_list.xlsx')
   const sheet= wb.Sheets['Sheet1']
   const data= xlsx.utils.sheet_to_json(sheet);
   
   // Making API calls for fetching the product price
   try {
    await Promise.all(data.map(async (i) => {
    const product_code=i.product_code
    const price= await axios.get(`https://api.storerestapi.com/products/${product_code}`)
    var individual_product_price=price.data.data.price
    product_details.set(product_code,individual_product_price)
   }))
   }
   
   catch (error) {
    return res.json({error:error.message});
   }
    
   //Filling the prices for the products
    for (let index = 2; index <= data.length+1; index++) {
            let product_name = sheet[`A${index}`].v;
            let product_final_price=product_details.get(product_name)
            sheet[`B${index}`]={v:product_final_price}
        }
     xlsx.writeFile(wb,'./sheet/product_list.xlsx');

    // Download the sheet 
     res.download(path.resolve('./sheet/product_list.xlsx'));
     
  });

app.listen('4000',()=> console.log('Running'));
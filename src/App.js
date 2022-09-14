import './App.css';
import Axios from "axios"
import { useEffect, useState} from 'react'
import * as XLXS from "xlsx"
import ExcelJs from "exceljs";

function App() {
  const [getDataURL, setGetDataURL] = useState("");
  const [token, setToken] = useState("");
  const [oldUrl, setOldUrl] = useState([]);//舊的網址
  const [newUrl, setNewUrl] = useState([]);//獲得的網址
  const [fileName, setFileName] = useState(null);//excel文件
  const [upLoadFile, setUpLoadFile] = useState("");
  const [getfinish,setGetFinish] = useState(Boolean);//是否全部獲取完畢
  const [tagfinish,setTagFinish] = useState(Boolean);//是否全部獲取完畢
  const [encodeID, setEncodeID] = useState([])//encodeID
  const [tagName, setTagName] = useState("");
  const [tagid, setTagID] = useState([]);
  //處理文件上傳(excel檔)
  const handleFile = async (e) =>{
    const file = e.target.files[0];
    setFileName(file.name)

    const data = await file.arrayBuffer();
    let workurl = XLXS.read(data);
    const worksheet = workurl.Sheets[workurl.SheetNames[0]];
    
    const jsonData = XLXS.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
    });
    for(let i = 1; i < jsonData.length; i++){
      setOldUrl(data=>[...data,jsonData[i][0]]);
    }
  }
  //
  const changeToken = ()=>{
    console.log(token);
    setGetDataURL("https://api.pics.ee/v1/links/?access_token="+ token);
  }
  //得到縮短的網址
  const onClick = async ()=>{
    console.log(getDataURL);
    for(let i = 0; i < oldUrl.length; i++){
      let Data = await Axios.post(getDataURL, {
        "url" : oldUrl[i],
        "applyDomain": true,
      });
      let dataUrl = Data.data.data.picseeUrl;
      let encode = dataUrl.split("/");
      console.log(encode[3]);

      setEncodeID(data=>[...data, encode[3]]);
      setNewUrl(data=>[...data,dataUrl]);
      console.log(Data);
    }
    
    
    setGetFinish(true);
  }
  //
  const addTag = async () =>{
    for(let i = 0; i < encodeID.length; i++){
      let Data = await Axios.post(`https://api.pics.ee/v1/links/`+encodeID[i]+`/tags?access_token=` + token,
      {"value": tagName
      }
      );
      console.log(encodeID[i]);
      console.log(Data);
      let tagID = Data.data.data.id;
      console.log(tagID);
      setTagID(data=>[...data, tagID]);
      console.log(tagName);
    }

    setTagFinish(true);
  }
  


  //轉換成excel檔
  function changeExcel(){
    const workbook = new ExcelJs.Workbook(); // 創建試算表檔案
    const sheet = workbook.addWorksheet('工作表範例1'); //在檔案中新增工作表 參數放自訂名稱
    let row = [];
    for(let i = 0; i < oldUrl.length; i++){
      row.push([oldUrl[i],newUrl[i],tagName,tagid[i]]);
    }
    console.log(row);
		sheet.addTable({ // 在工作表裡面指定位置、格式並用columsn與rows屬性填寫內容
	    name: 'table名稱',  // 表格內看不到的，讓你之後想要針對這個table去做額外設定的時候，可以指定到這個table
	    ref: 'A1', // 從A1開始
	    columns: [{name:'原本的網址'},{name:'新的網址'},{name:'tagName'},{name: 'tagID'}],
	    rows: row
		});
    //改變表格樣式
    sheet.getColumn(1).width = 90;
    sheet.getColumn(2).width = 50;

    // 表格裡面的資料都填寫完成之後，訂出下載的callback function
		// 異步的等待他處理完之後，創建url與連結，觸發下載
	  workbook.xlsx.writeBuffer().then((content) => {
		const link = document.createElement("a");
	    const blobData = new Blob([content], {
	      type: "application/vnd.ms-excel;charset=utf-8;"
	    });
	    link.download = upLoadFile +'.xlsx';
	    link.href = URL.createObjectURL(blobData);
	    link.click();
	  });
	}


  return (
    <div className="App">
      <h1>問卷調查處理</h1>
      
      <form className='upload'>
      <h1>轉換網址</h1>
        <label className='uploadFile'>
          上傳檔案
              <input type={"file"} onChange={(e)=> handleFile(e)}/>
        </label> {fileName}
              {/* {oldUrl.map(item=><li>{item}</li>)} */}
      </form>
      <div className='getData'>
        <button onClick={(e)=>changeToken()}>SetToken</button>
        {/* {getDataURL} */}
        <input type="text" placeholder="access_token" onChange={(e)=>setToken(e.target.value)}/>
      </div>
      <div className='getData'>
        <button onClick={(e)=>onClick()}>getData</button> 
        {/* {newUrl.map(item=><li>{item}</li>)} */}
        {getfinish ? "GetData": "Dataloading"}
        <div>
        <button onClick={(e)=>addTag()}>getTag</button> 
          <input type="text" placeholder="tagName" onChange={(e)=>setTagName(e.target.value)}/>
          {tagfinish ? "AddTag": "Tagloading"}
        </div>
      </div>
      <div className='getData'>
        
        {/* <button onClick={(e)=>addTag()}>addTag</button>  */}
        {/* {encodeID.map(item=><li>{item}</li>)} */}
      </div>
      <div className='getData'>
        <label htmlFor="">檔名:</label>
        <input type="text" onChange = {(e)=>{setUpLoadFile(e.target.value)}}/>
        <button onClick={(e)=>changeExcel()}>轉換成Excel表</button>
      </div>
    </div>
  );
}

export default App;

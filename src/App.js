// App.js
import React from 'react';
import { tempData } from './tempConstant'; 
import ExportToExcel from './ExportToExcel';



function App() {


  return (
    <>
   <ExportToExcel tempdata={tempData} fileName={"ExcelExport"}/>
    </>
  );
};

export default App;

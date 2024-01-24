import React from "react";
import * as XLSX from "xlsx";

const ExportToExcel = ({ tempdata}) => {

  const calculateColumnWidths = (data) => {
   
    const widths = {};
    console.log(widths)
    data.forEach((row) => {
      row.forEach((cell, columnIndex) => {
        const cellWidth = cell.toString().length;
        if (!widths[columnIndex] || cellWidth > widths[columnIndex]) {
          widths[columnIndex] = cellWidth;
        }
      });
    });
    return Object.values(widths);
  };

  
  const exportToExcel = () => {
    const data = [tempdata];
    const req = data.map((value) => {
      return value.map((userValue) =>
        userValue.map((userData) => {
          const userKeys = Object.keys(userData);
          const user = userKeys[0];
          let userMappedData = { user };
    
          userData[user].forEach((problemData) => {
            const problemKey = [
              `Problem: ${problemData.Problem} & Iteration: ${problemData.Iteration} `,
            ];
    
            Object.keys(problemData).forEach((problemDataKeys) => {
              if (typeof userMappedData[problemKey] !== "object") {
                userMappedData[problemKey] = {};
              }
              userMappedData[problemKey][problemDataKeys] =
                problemData[problemDataKeys];
            });
          });
    
          return userMappedData;
        })
      );
    });
    

    const buildExcelData = (data) => {
      const excelData = [];
     
      const DataProblems = Array.from(
        new Set(
          data.flatMap((user) =>
            Object.keys(user).filter((key) => key.includes("Problem"))
          )
        )
      );
   
      // Header Row
      const headerRow = [""];
      DataProblems.forEach((problem) => {
        for (let i = 0; i < 5; i++) {
          headerRow.push(problem);
        }
      });
      
    
      
      excelData.push(headerRow);


      const subheaderRow = ["Users"];
      DataProblems.forEach(() => {
        const subheaders = [ "Decision","Rating", "Start Time", "End Time", "Email Response"];
        subheaderRow.push(...subheaders);
      });
      
      excelData.push(subheaderRow);

      // Data rows
      data.forEach((data) => {
        const userData = [data.user];
        DataProblems.forEach((problem) => {
          const problemKey = `${problem}`;
          const iterationData = data[problemKey];
          if (iterationData) {
            const {  decisions,rating, startTime, endTime, emailResponse } =iterationData;
            const formattedRating = Array.isArray(rating)? rating.join(","): "";
            const iterationValues = [];
            iterationValues[0] = decisions || "";
            iterationValues[1] = formattedRating || "";
            iterationValues[2] = startTime || "";
            iterationValues[3] = endTime || "";
            iterationValues[4] = emailResponse || "";
            userData.push(...iterationValues);
          } 
        });

        excelData.push(userData);
      });
      console.log(excelData);

      return excelData;
    };

      // Generate Excel data
      const excelData = buildExcelData(req.flat(Infinity));
      const columnWidths = calculateColumnWidths(excelData);
      const ws = XLSX.utils.aoa_to_sheet(excelData);
      ws['!cols'] = columnWidths.map((width) => ({ width: width * 1}));

    const numberOfRepeats = 16;
    const merge = [];
    const columnsPerIteration = 5;
   
    
    for (let i = 0; i < numberOfRepeats; i++) {
      const startColumn = 1 + i * columnsPerIteration;
      const endColumn = startColumn + columnsPerIteration - 1;
    
      const entry = {
        s: { r: 0, c: startColumn },
        e: { r: 0, c: endColumn },
      };
    
      merge.push(entry);
    }
    
    ws["!merges"] = merge;
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
    const sheet = wb.Sheets["Sheet 1"];
    
    XLSX.writeFile(wb, "exported_data.xlsx");
    

  };

  return <button onClick={exportToExcel}>Export to Excel</button>;
};

export default ExportToExcel;

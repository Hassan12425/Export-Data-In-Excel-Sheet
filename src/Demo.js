import React from 'react';
import * as XLSX from 'xlsx';
import {data} from './Constant'

const Demo = () => {
  const exportToExcel = () => {

    const ws = XLSX.utils.json_to_sheet(buildExcelData(data));
    const numberOfRepeats = 16; // Adjust this based on the number of desired iterations
    const merge = [];
    const columnsPerIteration = 5;
    
    for (let i = 0; i < numberOfRepeats; i++) {
      const startColumn = 1 + i * columnsPerIteration;
      const endColumn = startColumn + columnsPerIteration - 1;
    
      const entry = {
        s: { r: 1, c: startColumn },
        e: { r: 1, c: endColumn },
      };
    
      merge.push(entry);
    }
    ws["!merges"] = merge
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');
    XLSX.writeFile(wb, 'exported_data.xlsx');
  };

  const buildExcelData = (data) => {
    const excelData = [];
  
    // Extract unique problems from the data
    const uniqueProblems = Array.from(
      new Set(data.flatMap((user) =>
        Object.keys(user).filter((key) => key.includes('Problem'))
      ))
    );
  
    // Header row
    const headerRow = [
      '',
      ...uniqueProblems.flatMap((problem) => [
        problem,
        problem,
        problem,
        problem,
        problem,
      ]),
    ];
    excelData.push(headerRow);
  
    // Subheader row
    const subheaderRow = ['Users', ...uniqueProblems.flatMap(() => ['Rating', 'Decision', 'Start Time', 'End Time', 'Email Response'])];
    excelData.push(subheaderRow);
  
    // Data rows
    data.forEach((user) => {
      const userData = [user.Users];
  
      // Iterate over problems and iterations
      uniqueProblems.forEach((problem) => {
        
        const problemKey = `${problem}`;
        const iterationData = user[problemKey];
        if (iterationData && iterationData[0]) {
          const iteration = iterationData;                    
          userData.push(iteration[1].Rating || '');
          userData.push(iteration[0].Decisions || '');
          userData.push(iteration[3]['Iteration Start Time'] || '');
          userData.push(iteration[4]['Iteration End Time'] || '');
          userData.push(iteration[2]['Email Response'] || '');
        } else {
          // If data is not available for the iteration, push empty values
          userData.push('', '', '', '', '');
        }
      });
  
      excelData.push(userData);
    });
  console.log(excelData);
  
    return excelData;
  };

  return (
    <div>
      <button onClick={exportToExcel}>Export to Excel</button>
    </div>
  );
};

export default Demo;
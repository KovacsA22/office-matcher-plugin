/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("select-delimiter").onclick = () => load("delimiter-cell");
    document.getElementById("select-texts").onclick = () => load("texts-start");
    document.getElementById("select-results").onclick = () => load("results-start");
    document.getElementById("count-tags").onclick = runCountTags;
    document.getElementById("create-histogram").onclick = runCreateHistogram;
  }
});

let delimiter="", textArray=[], resultAddress;

function getRangeById(context, id){
  return context.workbook.worksheets.getActiveWorksheet().getRange(document.getElementById(id).value);
}

async function load(targetId){
  try {
    await Excel.run(async (context) => {
      let range = context.workbook.getSelectedRange();
      
      // Read the range address
      range.load("address");
      await context.sync();
      document.getElementById(targetId).value=range.address;
      
    });
  } catch (error) {
    console.error(error);
  }
}

async function processInputs(context){
    // DELIMITER
    {
      var delimiterRange = getRangeById(context, "delimiter-cell");
      //let delimiterRange = context.workbook.worksheets.getActiveWorksheet().getRange(document.getElementById("delimiter-cell").value.toString());
      //let delimiterRange=context.workbook.worksheets.getItem("Sheet1").getRange("G4:G4");
      delimiterRange.load("values");
      await context.sync();

      
      delimiter=delimiterRange.values[0][0];
      document.getElementById("delimiter").innerText = "\""+delimiter+"\"";
    }

    // TEXTS
    {
      let textRange = getRangeById(context, "texts-start");

      textRange.load("columnIndex");
      await context.sync();

      let top = context.workbook.worksheets.getActiveWorksheet().getCell(0,textRange.columnIndex);

      top.load("values");
      await context.sync();

      const textCount = top.values[0][0];
      document.getElementById("text-count").innerText = textCount + " db";

      textRange = textRange.getResizedRange(textCount-1,0);
      textRange.select();

      textRange.load("values");

      await context.sync();
      textArray=[];
      for(const text of textRange.values){
        textArray.push(text[0].toString().toLowerCase());
      }
    }

    //RESULTS
    resultAddress = document.getElementById("results-start").value;
}


async function runCountTags() {
  try {
    await Excel.run(async (context) => {

      await processInputs(context);

      const occurences = [];
      for(const text of textArray){
        const array = text.split(delimiter);
        if(array.length==1 && array[0]==""){
          occurences.push(0);
        }else{
          const occurence = text.split(delimiter).length;
          occurences.push(occurence);
        }
      }

      console.log(occurences);

      let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      resultRange=resultRange.getResizedRange(occurences.length-1,0);
      resultRange.select();
      resultRange.values=[];
      for(const occurence of occurences){
        resultRange.values.push([occurence]);
      }

    });
  } catch (error) {
    console.error(error);
  }
}


async function runCreateHistogram() {
  try {
    await Excel.run(async (context) => {
      await processInputs(context);

      let maxTagCount=0;
      const tagCounts = {};
      for(const text of textArray){
        const array = text.split(delimiter);
        let tagCount=0;
        if(array.length !=1 || array[0] != ""){
          tagCount = text.split(delimiter).length;
        }
        if(tagCount>maxTagCount){
          maxTagCount=tagCount;
        }
        if(tagCounts[tagCount+""] == undefined){
          tagCounts[tagCount+""] = 0;
        }
        tagCounts[tagCount+""] += 1;
        
      }

      console.log(tagCounts);

      let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      resultRange=resultRange.getResizedRange(maxTagCount,1);
      resultRange.select();
      resultRange.values=[];
      for(i=0;i<=maxTagCount;i++){
        resultRange.values.push([i,tagCounts[i+""]==undefined?0:tagCounts[i+""]]);
      }

      
    });
  } catch (error) {
    console.error(error);
  }
}


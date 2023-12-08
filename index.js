
// for this purpose we are going to use xlsx npm package
import xlsx from "xlsx";
import jsonObject from "./data.json" assert { type: "json" };

let myBook = xlsx.utils.book_new();

let sheets = [];//this will store all the new sheets with the their name
let changedJson = {};//this will store new modified object
let toBeConverted = [];// this is passed to json to sheet finally after the object is modified

console.log(Array.isArray([jsonObject]))

function newData(jsonObject) {
  for (const row in jsonObject) {
    // ERROR Handling if its null pass EMPTY as its value
    if (jsonObject[row] === null) {
      changedJson[row] = "EMPTY";
      continue;
    }

    //checking if it is an array
    if (jsonObject[row].constructor === [].constructor) {

    //checking if it is an array of Object
      if (jsonObject[row][0].constructor === {}.constructor) {

    // making new sheet as it is array of Object now
        changedJson[row] = `=HYPERLINK("#${row}!A1", "SHEET::${row}")`;
        let mySheet = xlsx.utils.json_to_sheet(jsonObject[row]);
        sheets.push({ mySheet, name:row });
      } 

    // if its array but not of object hence return its element only
        else {
        changedJson[row] = jsonObject[row].join(" ");
      }
    }
        else if (jsonObject[row].constructor === {}.constructor) {
        // making newSheet again if there is object literals
        // and linking new sheets
        changedJson[row] = `=HYPERLINK("#${row}!A1", "SHEET::${row}")`;
        let mySheet = xlsx.utils.json_to_sheet([jsonObject[row]]);
        sheets.push({ mySheet, name:row });
        } 
    else {
      changedJson[row] = jsonObject[row];
    }
  }
}

if(Array.isArray(jsonObject))
newData(jsonObject[0]);//if json file is array then pass it as object
else
newData(jsonObject)//if json is object literal then pass it as it is
toBeConverted = [changedJson];
let mySheet = xlsx.utils.json_to_sheet(toBeConverted);
xlsx.utils.book_append_sheet(myBook, mySheet);
for (let i = 0; i < sheets.length; i++) {
  xlsx.utils.book_append_sheet(myBook, sheets[i].mySheet, sheets[i].name);
}
xlsx.writeFile(myBook, "convertedJsonToExcel.xlsx");








// Note: Due to my exams are running on i wanted to make it more precise and add some functionality, what i learnt in one day i have put it all in the above code.
// 9118372331 ,goluojha13101992@gmail.com

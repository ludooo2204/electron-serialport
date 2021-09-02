const {shell} = require('electron');
// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.

const serialport = require("serialport");
const tableify = require("tableify");
const Readline = require("@serialport/parser-readline");
async function listSerialPorts() {
	await serialport.list().then((ports, err) => {
		if (err) {
			document.getElementById("error").textContent = err.message;
			return;
		} else {
			document.getElementById("error").textContent = "";
		}
		console.log("ports", ports);
		if (ports.length === 0) {
			document.getElementById("error").textContent = "No ports discovered";
		}

		tableHTML = tableify(ports);
		document.getElementById("ports").innerHTML = tableHTML;
	});
}
listSerialPorts();
port = new serialport("COM11", { baudRate: 115200 });
// Set a timeout that will check for new serialPorts every 2 seconds.
// This timeout reschedules itself.
// setTimeout(function listPorts() {
//   listSerialPorts();
//   setTimeout(listPorts, 2000);
// }, 2000);


if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('fichierTransfert.xls');
/* DO SOMETHING WITH workbook HERE */

// get first sheet
let first_sheet_name = workbook.SheetNames[0];
let worksheet = workbook.Sheets[first_sheet_name];

// read value in D4 
let cell = worksheet['A1'].v;
console.log(cell)



// write to new file
// formatting from OLD file will be lost!


const meas = document.getElementById("click");
const meas2 = document.getElementById("click2");
const save = document.getElementById("click3");
const write = document.getElementById("click4");
const input = document.getElementById("inputValue");
const inputForm = document.getElementById("inputForm");
inputForm.addEventListener("submit", (e) => { 
	e.preventDefault()
	console.log(input.value)
	port.write("SOUR "+input.value+"\n");
	// console.log(e.target.value)
})
meas.addEventListener("click", () => {
	console.log("rem");
	port.write("rem\n");
});
save.addEventListener("click", () => {
	
	// modify value if A3 is undefined / does not exists
	// XLSX.utils.sheet_add_aoa(worksheet, [['NEW VALUE from NODE']], {origin: 'A3'});
	
	worksheet['B1'].v =4 ;
	XLSX.writeFile(workbook, 'fichierTransfert.xls');
});
write.addEventListener("click", () => {

	// worksheet['D16'].v = 1666;
// Open a local file in the default app
shell.openPath('C:\\Users\\ludovic.vachon\\electron\\serial\\electron-serialport\\vide v2.0.xls');

// Open a URL in the default way
// shell.openExternal('https://github.com');
	
});
// port.write("rem",(toto)=>console.log(toto))
meas2.addEventListener("click", () => {
	// console.log("meas?");
	// port.write("meas?\n");
	// port.write("SOUR:Func TC\n");
	port.write("SOUR 266\n");
	
});
//rem
//SOUR:Func TC
//'SOUR:tcouple:type ' + voie.typeTc);
//SOUR 1200
console.log(port);
port.on("error", function (err) {
	console.log("error!");
	console.log("Error: ", err.message);
});

// Read data that is available but keep the stream in "paused mode"
// port.on('readable', function () {
//   console.log("read")
//   console.log('Data:', port.read())
// })

// Switches the port into "flowing mode"
// port.on('data', function (data) {
//   console.log('Data:', data)
//   console.log(data.toString('utf8'));
// })

// Pipe the data into another stream (like a parser or standard out)
const lineStream = port.pipe(new Readline());
console.log(lineStream);
lineStream.on("data", (data) => console.log(data));




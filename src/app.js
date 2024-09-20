const fs = require("fs");
const docx = require("docx");
const { exec } = require("child_process");
const path = require("path");
//const { createCanvas } = require("canvas");
const jsBarcode = require("jsbarcode");

const { spawn } = require("child_process");

let command = `Start-Process "G:\\Programiranje\\JavaScript\\Electron js\\zebra-printer-app\\word_docx\\output.docx" -Verb print`;

const sku1_field = document.getElementById("sku1");
const sku2_field = document.getElementById("sku2");
const sku3_field = document.getElementById("sku3");

const printButton = document.getElementById("printBtn");
const ressetBarcode = document.getElementById("ressetBarcode");

const barcodeImgField = document.getElementById("barcodeImg");

const generateBarcode = () => {
    let value = "";
    for (let i = 0; i < 12; i++) {
        value += Math.floor(Math.random() * 9)
    } 

    console.log(value);
    return value;
}

let barcodeImage;

function generateBarcodeImage() {
    let barcodeVal = generateBarcode();

    console.log(barcodeVal);

    jsBarcode(barcodeImgField, barcodeVal, {
        format: 'CODE39', // Barcode format
        width: 2, // Width of each bar
        height: 30, // Height of the barcode
        displayValue: true, // Show the value below the barcode
        fontSize: 15,
    });
}

window.onload = () => {
    generateBarcodeImage();
}
ressetBarcode.addEventListener("click", generateBarcodeImage);

printButton.addEventListener("click", () => {
    let barCodeImg = new docx.ImageRun({
        type: "png",
        data: barcodeImgField.src,
        transformation: {
            width: 113,
            height: 30
        }
    })
    
    let doc = new docx.Document({
        sections: [
            {
                properties: {page: {size: {width: "3cm", height: "2cm"}, margin: {left: "0cm", right: "0cm", top: "0cm", bottom: "0cm"}}},
                children: [
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        border: {
                            bottom: {
                                style: docx.BorderStyle.SINGLE,
                                size: 3,
                                color: "000000" 
                            }
                        },
                        children: [
                            new docx.TextRun({
                                text: "S K U",
                                size: 20,
                                bold: true,
                                font: "Calibri"
                            }),
                        ]
                    }),
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [
                            new docx.TextRun({
                                text: "1                           2                           3",
                                size: 10,
                                bold: true,
                                font: "Calibri",
                            })
                        ]
                    }),
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        border: {
                            top: {
                                style: docx.BorderStyle.SINGLE,
                                size: 3,
                                color: "000000"
                            },
                            bottom: {
                                style: docx.BorderStyle.SINGLE,
                                size: 3,
                                color: "000000" 
                            }
                        },
                        children: [
                            new docx.TextRun({
                                text: `${sku1_field.value}  |  ${sku2_field.value}  |  ${sku3_field.value}`,
                                size: 17,
                                bold: true,
                                font: "Calibri",
                            })
                        ]
                    }),
                    new docx.Paragraph({
                        children: [
                            barCodeImg
                        ]
                    })
                ],
            },
        ],
    });
    
    docx.Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("./word_docx/output.docx", buffer);
        console.log("DOCX file created successfully!");
    });
    
    setTimeout(() => {
        spawn("powershell.exe", [command]);

        sku1_field.value = "";
        sku2_field.value = "";
        sku3_field.value = "";

        generateBarcodeImage();
    }, 1000);
})

//const canvas = createCanvas(113, 30);





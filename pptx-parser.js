import fs from "fs/promises";
import JSZip from "jszip";
import { parseStringPromise } from "xml2js";

//const regex = /<a:t>(.*?)<\/a:t>/gs;
const xmlFiles = [];
const data = fs.readFile("./test.pptx");
const zip = await JSZip.loadAsync(data);
const slides = zip.folder("ppt/slides");

slides.forEach(async slideFileName => {
    if(slideFileName.endsWith('.xml')){
        xmlFiles.push(slideFileName);
    }
});

for (const slide of xmlFiles) {
    const xmlData = await zip.file(`ppt/slides/${slide}`).async('string');
    const result = await parseStringPromise(xmlData);
    const formatResult = result['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'];
    if (!formatResult?.[0]) continue;

    let largestFontSize = 0;
    let title = '';
    let body = [];

    for (const shape of formatResult) {
        const shapes = shape['p:txBody'];
        for (const line of shapes) {
            const lines = line['a:p'];    
            for (const nestedLines of lines) {
                // console.log(nestedLines['a:r'][0]['a:t'][0]);
                const text = nestedLines['a:r'][0]['a:t'][0];
                const fontSize = nestedLines['a:r'][0]['a:rPr']?.[0]['$'].sz;
                if (fontSize > largestFontSize) {
                    largestFontSize = fontSize;
                    title = text;
                } else if (text != '') {
                    body.push(text.toString().trim());
                }
                
            }
        }
    }

    console.log(slide, body, title);
}



//console.log(formatResult);
/*
if (slideFileName.endsWith('.xml')) {
    xmlData = await zip.file(`ppt/slides/${slideFileName}`).async('string');
    const result = await parseStringPromise(xmlData);
    console.log(result);
    const formatResult = result['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'][0];
}        */

//console.log(pages);

//console.log(result['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'][0]);
/*
fs.readFile("./test.pptx", (err, data) => {
    if (err) throw err;

    JSZip.loadAsync(data).then((zip) => {
        const slideFiles = Object.keys(zip.files)
            .filter((filename) => filename.startsWith('ppt/slides/') && filename.endsWith('.xml'))
            .sort((a, b) => {
                const aSlideNumber = parseInt(a.match(/slide(\d+).xml/)[1]);
                const bSlideNumber = parseInt(b.match(/slide(\d+).xml/)[1]);
                return aSlideNumber - bSlideNumber;
            });
        console.log(slideFiles);

        // .sort() did not work
        /*
        Promise.all(
            slideFiles.map((filename) => {
                return zip.file(filename).async('string');
            })
        )    
        zip.file(filename).async('string')
            .then((text) => {
                const matches = text.match(regex);
                console.log(matches);
            })
            .catch((error) => {
                console.error(error);
            });
    });
});
*/
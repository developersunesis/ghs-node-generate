const fs = require('fs')
const PPT_Template = require('ppt-template');
const Presentation = PPT_Template.Presentation;
const Slide = PPT_Template.Slide;

const TEMPLATE = './templates/template1.pptx';
const OUTPUT = './tests/output.pptx';

// Presentation Object
var presentation = new Presentation();
console.log(`# Loaded ${TEMPLATE} as template`);

// Delete any pre-existing output 
if (fs.existsSync(OUTPUT)) {
    fs.unlink(OUTPUT, (err) => {
        if (err) throw err;
    });
}

// Load example.pptx
presentation.loadFile(TEMPLATE)
    .then(() => console.log('- Read Presentation File Successfully!'))
    .then(() => {
        var titlePageIndex = 1;
        var contentPageIndex = 2;

        // Get and clone slide. (Watch out index...)
        let titleSlide = presentation.getSlide(titlePageIndex).clone();
        let contentSlide = presentation.getSlide(contentPageIndex).clone();

        // Fill all content
        titleSlide.fillAll([
            Slide.pair('[Title]', 'Hello PPT'),
            Slide.pair('[Number]', 'this is a sample'),
        ]);

        contentSlide.fillAll([
            Slide.pair('[Content]', 'Some hymn content'),
        ]);

        var newSlides = [titleSlide, contentSlide];
        return presentation.generate(newSlides);
    }).then((newPresentation) => {
        console.log('- Generate Presentation Successfully');
        return newPresentation;
    }).then((newPresentation) => newPresentation.saveAs(OUTPUT))
    .then(() => console.log(`- ${TEMPLATE}, saved Successfully`))
    .catch((err) => console.error(err));
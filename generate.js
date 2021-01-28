const fs = require('fs')
var xml2js = require('xml2js');
const PPT_Template = require('ppt-template');

const Presentation = PPT_Template.Presentation;
const Slide = PPT_Template.Slide;

var presentation = new Presentation();
var parser = new xml2js.Parser();


fs.readFile('./data/hymn.txt', function (err, data) {
    parser.parseString(data, function (err, result) {
        var hymns = result
        hymns = hymns.Hymns.Hymn

        var length = hymns.length
        var i = 0;
        const interval = setInterval(() => {
            var rand = Math.floor(Math.random()*4)
            rand = rand <= 0 ? 1 : rand
            const TEMPLATE = './templates/template' + rand + '.pptx';

            const hymn = hymns[i]
            const number = hymn.id[0]
            const title = hymn.title[0]
            const stanzas = hymn.stanza
            const choruses = hymn.chorus
            const OUTPUT = `./hymns/${number.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '')} - ${title.replace(/[&\/\\#,+()$~%.'":*?<>{}]/g, '').trim()}.pptx`;

            if (fs.existsSync(OUTPUT)) {
                fs.unlink(OUTPUT, (err) => {
                    if (err) throw err;
                });
            }

            presentation.loadFile(TEMPLATE)
                .then(() => console.log(`- Prepared template for :: ${title} - ${number}`))
                .then(() => {
                    var titlePageIndex = 1;
                    var contentPageIndex = 2;
                    let titleSlide = presentation.getSlide(titlePageIndex).clone();

                    titleSlide.fillAll([
                        Slide.pair('[Title]', title.trim()),
                        Slide.pair('[Number]', `${number}`),
                    ]);

                    var contents = []
                    for (var j = 0; j < stanzas.length; j++) {
                        const stanza = stanzas[j]
                        let contentSlide = presentation.getSlide(contentPageIndex).clone();
                        contentSlide.fillAll([
                            Slide.pair('[Content]', stanza.replace('\r\n', '').trim()),
                        ]);
                        contents[contents.length] = contentSlide

                        if (choruses !== undefined) {
                            var chorus = choruses[j]
                            let contentSlide2 = presentation.getSlide(contentPageIndex).clone();
                            contentSlide2.fillAll([
                                Slide.pair('[Content]', chorus.replace('\r\n', '').trim()),
                            ]);
                            contents[contents.length] = contentSlide2
                        }
                    }

                    var newSlides = [titleSlide, ...contents];
                    return presentation.generate(newSlides);
                }).then((newPresentation) => {
                    console.log(`- Generate Presentation Successfully for :: ${title} - ${number}`);
                    return newPresentation;
                }).then((newPresentation) => newPresentation.saveAs(OUTPUT))
                .then(() => console.log(`- ${OUTPUT}, saved Successfully`))
                .catch((err) => console.error(err));

                i += 1;
                if(i >= length) clearInterval(interval)
        }, 10000)
    });
});
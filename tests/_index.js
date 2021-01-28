const PPTX = require('nodejs-pptx');
const fs = require('fs')

const pptx = new PPTX.Composer();

// pptx.load(`./templates/template1_.pptx`);
pptx.compose(presentation => {
  presentation.addSlide(slide => {
    slide.addText(text => {
      text.value('Hello World');
    });
  });
});

const filename = `./tests/hello-world.pptx`
if (fs.existsSync(filename)) {
    fs.unlink(filename, (err) => {
        if (err) throw err;
        console.log('Deleted!');
    });
}
pptx.save(filename);
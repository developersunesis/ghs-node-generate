var fs = require('fs'),
    xml2js = require('xml2js');
var parser = new xml2js.Parser();
var hymns = {}
fs.readFile('./data/hymn.txt', function (err, data) {
    parser.parseString(data, function (err, result) {
        hymns = result

        hymns = hymns.Hymns.Hymn
        for (var i = 0; i < hymns.length; i++) {
            const hymn = hymns[i]
            const number = hymn.id[0]
            const title = hymn.title[0]
            const stanzas = hymn.stanza
            const choruses = hymn.chorus

            for(var j = 0; j < stanzas.length; j++){
                const stanza = stanzas[j]
                var chorus = undefined
                if(choruses !== undefined){
                    chorus = choruses[j]
                }

                console.log(stanza)
                if(chorus !== undefined) 
                    console.log(chorus)
            }
        }
    });
});
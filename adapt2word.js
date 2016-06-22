var _ = require("underscore");
var async = require("async");
var fs = require("fs");
var mu = require("mu2");
var path = require("path");
var sizeOf = require('image-size');

var htmlTitle;
var htmlBody;
var transform = { tag: "div", class: "element" };



/**
 * Custom mutant method like Array.push(), but will check if the second item of the array exists before applying
 */
Array.prototype.pushInto = function(item) {
    if (item[1] !== undefined) this.push(item);
};



/**
 * Init
 */
function main() {
	console.log("Running adapt2word...");

	fs.readdir(".", function (err, files) {
		if (err) return console.log(err.toString());

		async.each(files, processFile, function() {
            console.log("Finished.");
        });
	});
}



/**
 * Process the metadata.json file from the AT export function
 */
function processFile(file, done) {
    if (file === 'metadata.json') {
    	fs.readFile(file, function (err, output) {
    		if (err) return console.log(err.toString());

    		htmlTitle = "";
    		htmlBody = "";
    		output = JSON.parse(output);

    		if (output.constructor !== Array) {
                convertToHTML(output);
            }
    		else for (var i = 0, j = output.length; i < j; i++) {
    			convertToHTML(output[i]);
    		}
            //convertToHTML(output);
    		createDocument(done);
    	});
    }
}



/**
 * Convert the data into html
 */
function convertToHTML(data) {

    htmlTitle = data.course.course[0].title;

    htmlBody += buildTable(
        'course',
        data.course.course[0].title,
        [
            ['Display title', data.course.course[0].displayTitle, ''],
            ['Body', data.course.course[0].body, '']
        ]
    );

    data.course.contentobject.forEach(function(co, coi, coa) {

        htmlBody += buildTable(
            'topic',
            co.title,
            [
                ['Display title', co.displayTitle, ''],
                ['Body', co.body, ''],
                ['Duration', co.duration, '']
            ]
        );

        data.course.article.forEach(function(a, ai, aa) {

            if (a._parentId !== co._id) return;

            // htmlBody += buildTable(
            //     'article',
            //     a.title,
            //     [
            //         ['Display title', a.displayTitle, ''],
            //         ['Body', a.body, '']
            //     ]
            // );

            data.course.block.forEach(function(b, bi, ba) {

                if (b._parentId !== a._id) return;

                // htmlBody += buildTable(
                //     'block',
                //     b.title,
                //     [
                //         ['Display title', b.displayTitle, ''],
                //         ['Body', b.body, '']
                //     ]
                // );

                data.course.component.forEach(function(c, ci, ca) {

                    if (c._parentId !== b._id) return;

                    //all
                    var rows = [
                        ['Component type', componentName(c._component), ''],
                        ['Display title', c.displayTitle, ''],
                        ['Body', c.body, '']
                        // ['Layout', c._layout, '']
                    ];

                    //mcq
                    rows.pushInto(['Instruction', c.properties.instruction, '']);
                    rows.pushInto(['Attempts', c.properties._attemps, '']);

                    if (c.properties._items) {
                        c.properties._items.forEach(function(p, pi) {
                            //mcq
                            rows.pushInto(['Item '+ (pi + 1), p.text, (p._shouldBeSelected ? 'Correct' : '')]);

                            if (p.feedback !== '') {
                                rows.pushInto(['Item '+ (pi + 1) +' feedback', p.feedback, '']);
                            }

                            //accordian
                            rows.pushInto(['Item '+ (pi + 1) +' title', p.title, '']);
                            rows.pushInto(['Item '+ (pi + 1) +' body', p.body, '']);
                            //only put in image details if they're inserted
                            if (p._graphic && p._graphic.src !== '') {
                                rows.pushInto(['Item '+ (pi + 1) +' graphic', processImage(p._graphic.src), '']);
                                rows.pushInto(['Item '+ (pi + 1) +' alt', p._graphic.alt, '']);
                            }

                            //matching
                            if (p._options) {
                                p._options.forEach(function(o, oi) {
                                    rows.pushInto(['Item '+ (pi + 1) +' option '+ (oi + 1), o.text, o._isCorrect ? 'Correct' : '']);
                                });
                            }
                        });
                    }

                    //mcq
                    if (c.properties._feedback) {
                        rows.pushInto(['Feedback correct', c.properties._feedback.correct, '']);
                        if (c.properties._feedback._incorrect) {
                            rows.pushInto(['Feedback incorrect final', c.properties._feedback._incorrect.final, '']);
                            rows.pushInto(['Feedback incorrect not final', c.properties._feedback._incorrect.notFinal, '']);
                        }
                        if (c.properties._feedback._partlyCorrect) {
                            rows.pushInto(['Feedback partly correct final', c.properties._feedback._partlyCorrect.final, '']);
                            rows.pushInto(['Feedback partly correct not final', c.properties._feedback._partlyCorrect.notFinal, '']);
                        }
                    }

                    //graphic
                    if (c.properties._graphic) {
                        rows.pushInto(['Image large', processImage(c.properties._graphic.large), '']);
                        rows.pushInto(['Image alt', c.properties._graphic.alt, '']);
                        // rows.pushInto(['Image attribution', c.properties._graphic.attribution, '']);
                    }

                    //slider
                    rows.pushInto(['Slider start label', c.properties.labelStart, '']);
                    rows.pushInto(['Slider end label', c.properties.labelEnd, '']);

                    if (c.properties._scaleStart) {
                        rows.push(['Slider range', c.properties._scaleStart +' - '+ c.properties._scaleEnd, '']);
                    }

                    if (c.properties._correctRange) {
                        rows.push(['Correct range', c.properties._correctRange._bottom +' - '+  c.properties._correctRange._top, '']);
                    }

                    htmlBody += buildTable(
                        'component',
                        c.title,
                        rows
                    );
                });
            });
        });
    });
}



/**
 * Returns a human readable name for a component type (would be better to get this from the json)
 */
function componentName(type) {
    switch (type) {
        case 'accordion':
            return 'Accordion';
        case 'graphic':
            return 'Graphic';
        case 'matching':
            return 'Matching Question';
        case 'mcq':
            return 'Multiple Choice Question';
        case 'slider':
            return 'Slider';
        case 'text':
            return 'Text';
    }
}



/**
 * Scale image down to a reasonable size
 */
function processImage(src, callback) {

    if (src) {
        var source = path.resolve(path.resolve(src).replace('course', ''));
        var size = sizeOf(source);
        var maxWidth = 200;
        var maxHeight = 300;
        var scale = Math.max(maxWidth / size.width, maxHeight / size.height);
        var scaledWidth = Math.round(size.width * scale);
        var scaledHeight = Math.round(size.height * scale);

        if (scaledWidth && scaledHeight) {
            return '<img src="'+ source +'" width="'+ scaledWidth +'" height="'+ scaledHeight +'" />';
        }
        else {
            return '<img src="'+ source +'" width="'+ maxWidth +'" height="'+ maxWidth +'" />';
        }
    }

    return undefined;
}



/**
 * Build an html table from an array
 *
 * @param  {string} type
 *         This determines the heading used and adds a class name to the table wrapper
 * @param  {string} title
 *         The title of the table
 * @param  {Array} rows
 *         An array of arrays of table data eg. [
 *         		['row 1 column 1', 'row 1 column 2'],
 *         		['row 2 column 1', 'row 2 column 2']
 *         ]
 * @return {string} The html string of the table
 */
function buildTable(type, title, rows) {

    switch (type) {
        case 'course': heading = 'h1';
            break;
        case 'topic': heading = 'h1';
            break;
        case 'article': heading = 'h2';
            break;
        case 'block': heading = 'h3';
            break;
        case 'component': heading = 'h4';
            break;
    }

    var output = '<div class="'+ type +'">';

    type = type.charAt(0).toUpperCase() + type.substr(1);

    output = '<'+ heading +'>'+ type +' - '+ title +'</'+ heading +'>';
    output += '<table>';

    if (rows) {
        output += '<tbody>';
        rows.forEach(function(row, ri) {
            var bg = ri % 2 ? '#FFFFFF' : '#EFEFEF';
            output += '<tr style="background:'+ bg +'">';
            row.forEach(function(td, tdi) {
                output += '<td style="';
                switch (tdi) {
                    case 0:
                        output += 'width:20%;font-weight:bold';
                        break;
                    case 1:
                        output += 'width:65%';
                }
                output += '"><div style="border:10px solid '+ bg +'">'+ td +'</div></td>';
            });
            output += '</tr>';
        });
        output += '</tbody>';
    }

    output += '</table></div><br /><br /><br /><br />';

    return output;
}



/**
 * Creates the word document from the html gathered above
 */
function createDocument(done) {

	var template = "";
	var filename = htmlTitle + ".doc";

	mu.compileAndRender(path.resolve(__dirname, "template.html"), {
		title: htmlTitle,
		body: htmlBody
	})
    .on("data", function(data) {
		template += data;
	})
    .on("end", function() {
        fs.writeFile(filename, template, function (err) {
            if (err) return console.log(err.toString());

            console.log("Written " + filename + ".");
            done();
        });
	});
}

module.exports = {
	main: main
};

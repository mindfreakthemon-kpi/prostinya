var RANK_REGEXP = /(\d+)(\w)(\d+)?/,
	
	options = require('./options'),
	fs = require('fs'),
	xlsx = require('node-xlsx'),
	swig  = require('swig'),
	Iconv  = require('iconv').Iconv,
	template = swig.compileFile('./template.val.tpl');

var iconv = new Iconv('UTF-8', 'CP1251');

var obj = xlsx.parse(options.input),
	list = obj[options.page | 0].data.map(function (row) {
		return row.splice(2)
	}),
	header = list.shift(),
	groupName = header.shift(),
	ranks = {},
	map = {};

// remove line with numbers	
list.shift();

list.forEach(function (row, rowId) {
	var id = row.shift();

	row.forEach(function (val, i, array) {
		if (val === '*') {
			array[i] = null;
		} else if (val) {
			array[i] = val.replace(/\s+/g, '');
		}
	});

	ranks[id] = row;
});

header.forEach(function (name, i) {
	map[name] = [];

	Object.keys(ranks)
		.forEach(function (key) {
			var str = stringifyRank(ranks[key][i]);
			map[name].push(str);
		});
});

function stringifyRank(rank) {
	if (!rank) {
		return '';
	}

	var matches = rank.match(RANK_REGEXP);

	return '' + matches[1] + ' ' + matches[2] + ' ' + ({
		'3': 'задовільно',
		'4': 'добре',
		'5': 'відмінно'
	}[matches[3]] || 'зараховано')
}

// console.log(groupName);
// console.log(JSON.stringify(header, null, '\t'));
// console.log(JSON.stringify(ranks[0], null, '\t'));
// console.log(JSON.stringify(map, null, ' '));

try {
	fs.mkdirSync('./' + groupName);
} catch (e) {}

Object.keys(map)
	.forEach(function (name) {
		var _names = name.split(' '),
			lastName = _names.shift(),
			firstName = _names.shift(),
			middleName = _names.shift();

		var output = template({
			ranks: map[name],
			lastName: lastName,
			firstName: firstName,
			middleName: middleName,
			name: middleName
		});

		fs.writeFileSync('./' + groupName + '/' + lastName + '.val', iconv.convert(output.trim()));
	});


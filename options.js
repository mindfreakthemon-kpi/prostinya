module.exports = require('nomnom')
	.option('input', {
		abbr: 'i',
		required: true,
		help: 'Input XLSX to process'
	})
	.option('page', {
		abbr: 'p',
		required: true,
		default: 0,
		help: 'Page number'
	})
	.parse();
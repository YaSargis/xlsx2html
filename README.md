# xlsx2html


## Getting started
### Module convert xlsx to html with saving styles 

#### Install:
`npm i xlsx2html`
#### Example:
```
const x2h = require('xlsx2html');
const fs = require('fs');

const fileData = fs.readFileSync('./test.xlsx');
x2h.xlsx2html(fileData).then( (html) => {
	// html use
	
})
```
## Authors and acknowledgment
Sargis Kazaryan
## License
MGLC
# msg2eml.js
A javascript library to convert Outlook *.msg files to *.eml natively in the browser.

## Demo
Try it out [here](https://master131.github.io/msg2eml.js/demo/index.html).

## Installation & Building
Make sure you have Typescript installed globally (or adjust the package.json accordingly).

``npm install``

``npm run build``

lib/dist/msg2eml.bundle.js (non-minified) and lib/dist/msg2eml.min.js (minified) will be generated.

This library can function in IE11 given the correct polyfills are present (see the demo) and the "use strict"/'use strict' keywords have been stripped from the msg2eml.js library.

## API
This adds a function called msg2eml into the global window object.

| Function Name  | Parameters                                                    | Returns                                                                             | Description             |
|----------------|---------------------------------------------------------------|-------------------------------------------------------------------------------------|-------------------------|
| window.msg2eml | Blob - a blob object containing the contents of the .msg file | Promise<string> - a promise which will generate the string content of the .eml file | Converts a .msg to .eml |

See the demo directory for an example of this in-use.

## Thanks
This would not be possible without all the hard work of the people who wrote a lot of the underlying libraries which msg2eml.js depends on:
- https://github.com/JoshData/convert-outlook-msg-file
- https://github.com/mazira/rtf-stream-parser
- https://github.com/SheetJS/js-cfb
- https://github.com/papnkukn/eml-format
- https://github.com/HiraokaHyperTools/DeCompressRTF
- https://github.com/moment/moment
- https://github.com/peterolson/BigInteger.js

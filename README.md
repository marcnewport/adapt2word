# adapt2word

Builds an MS Word document of an Adapt course from the "Export course" function in the authouring tool.

## Installation

Requires [Node.js](http://nodejs.org) to be installed.

From the command line, run:
```
npm install -g adapt2word
```

## Usage

1. In the authoring tool download the zip file using the "Export course" button from the course view
2. Unzip that package on your local machine
3. From the command line, cd into that directory
4. Run the `adapt2word` command
5. The word document will be placed in the current directory

If you want to embed the images into the document, open it in MS Word (2013) and then go:
* File -> Related documents (bottom right) -> Edit Links to files
* Select all items in the list (Shift+Left Mouse)
* Tick the box "Save picture in document"
* Update now -> OK
* Save the document as a docx

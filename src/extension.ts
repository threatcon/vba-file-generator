import * as vscode from "vscode";
const fs = require("fs");
const path = require("path");

export function activate(context: vscode.ExtensionContext) {
    let disposable = vscode.commands.registerCommand('extension.newVbaFile', async () => {

		if (vscode.window.activeTextEditor) {
			const folderPath = path.dirname(
			  vscode.window.activeTextEditor?.document.fileName
			  
			);
			console.log(folderPath);


        const fileName = await vscode.window.showInputBox({
            prompt: 'Enter a file name',
            placeHolder: 'e.g. MyFile'
        });

        if (fileName) {
            const fileExtension = await vscode.window.showQuickPick(['.cls', '.bas'], {
                placeHolder: 'Select a file extension'
            });

            if (fileExtension) {
				const fullFileName =  folderPath +'\\' + fileName + '.' + fileExtension;
				console.log(fullFileName);
                const fileContent = getFileContent(fileExtension, fileName);
				if (fs.existsSync(fullFileName) ) {
					vscode.window.showErrorMessage("ERROR! File Exists!");
					return;
				}
				fs.writeFileSync(fullFileName, fileContent);

				vscode.workspace.openTextDocument(vscode.Uri.file(fullFileName)).then(file => {
					vscode.window.showTextDocument(file);
				});
			}
            }
        }
});

	vscode.window.showInformationMessage('VBA New File Generator Started!');
    context.subscriptions.push(disposable);
}

function getFileContent(fileExtension: string, fileName: string): string {
    if (fileExtension === '.cls') {
        return `
VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "` + fileName + `"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
        `;
    } else if (fileExtension === '.bas') {
        return `
Attribute VB_Name = "` + fileName + `"

Option Explicit
`;
    }

    return '';
}

//const filePath = '/path/to/file.txt';
//const fileContent = 'This is the content of the file.';

//createFile(filePath, fileContent);

// This method is called when your extension is deactivated
export function deactivate() {}

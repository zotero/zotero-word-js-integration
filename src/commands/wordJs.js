/*
	***** BEGIN LICENSE BLOCK *****
	
	Copyright © 2023 Corporation for Digital Scholarship
                     Vienna, Virginia, USA
					http://zotero.org
	
	This file is part of Zotero.
	
	Zotero is free software: you can redistribute it and/or modify
	it under the terms of the GNU Affero General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	Zotero is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU Affero General Public License for more details.

	You should have received a copy of the GNU Affero General Public License
	along with Zotero.  If not, see <http://www.gnu.org/licenses/>.
	
	***** END LICENSE BLOCK *****
*/
/* global global, Office, self, window */
Office.onReady(() => {
	// If needed, Office.js is ready to be called
	OfficeExtension.config.extendedErrorLogging = true;
});

function generateButtonHandler(button) {
	return function (event) {
		Zotero.debug(`Clicked addin button ${button}`)
		let session = new Zotero.Session(event, button);
		session.execCommand(button);
	}
}

function getGlobal() {
	return typeof self !== "undefined"
		? self
		: typeof window !== "undefined"
			? window
			: typeof global !== "undefined"
				? global
				: undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
const buttons = [
	'addEditCitation',
	'addEditBibliography',
	'addNote',
	'citationExplorer',
	'setDocPrefs',
	'refresh',
	'unlink'
]
for (let button of buttons) {
	g[button] = generateButtonHandler(button);
}

function handleError(e) {
	let result = {
		error: e.type || `Connector Error`,
		message: e.message,
		stack: e.stack,
	}
	if (e.debugInfo) {
		result = {...result,
			errorLocation: e.debugInfo.errorLocation,
			fullStatements: e.debugInfo.fullStatements,
			surroundingStatements: e.debugInfo.surroundingStatements
		}
	}
	console.log(result);
	debugger
	throw e;
}

// Fixed
g.testFootnoteFieldRetrieve = async function(event) {
	await Word.run(async (context) => {
		const selection = context.document.getSelection();
		const footnote = selection.insertFootnote('test');
		await context.sync();
		const footnoteBodyRange = footnote.body.getRange();
		const field = footnoteBodyRange.insertField('End', 'Addin', 'Test Field');
		await context.sync();
		field.result.insertText("TEST FIELD AGAIN", "Replace");
		await context.sync();
		field.code = "ADDIN test";
		await context.sync();
		
		const fields = footnote.body.fields.load({ result: { text: true } });
		await context.sync();
		console.log(fields.items[0].result.text);
	});
	if (event) {
		event.completed();
	}
}

// Now works
g.testInsertHtmlIntoField = async function(event) {
	await Word.run(async (context) => {
		const selection = context.document.getSelection().getRange();
		const field = selection.insertField('Replace', 'Addin');
		field.code = `ADDIN test`;
		field.result.insertHtml("Test", "Replace");
		field.track();
		await context.sync();
	});
	if (event) {
		event.completed();
	}
}

// Broken on my account
g.testInsertField = async function(event) {
	try {
		await Word.run(async (context) => {
			const range = context.document.getSelection().getRange();
			range.insertField('Replace', 'Addin');
			await context.sync();
			// Throws RichApi.Error: Sorry, this function isn't available.
			// Perform a runtime check on the Office add-in to find out whether
			// the feature is supported by the host.
		});
	}
	catch (e) {
		console.log(e);
	}
	if (event) {
		event.completed();
	}
}

// Fixed
g.testFieldCodeChangeAfterHtml = async function(event) {
	try {
		await Word.run(async (context) => {
			const selection = context.document.getSelection().getRange();
			const field = selection.insertField('Start', 'Addin');
			field.code = `ADDIN test`;
			// field.result.insertText("Test1", "Start");
			field.result.insertHtml("<b>Test1</b>", "Start");
			await context.sync();
			field.code = `ADDIN test1after`
			// Throws a GeneralException
			await context.sync();
		});
	}
	catch (e) {
		handleError(e);
	}
	if (event) {
		event.completed();
	}
}

g.testTrackedFieldChange = async function(event) {
	try {
		let tracked = [];
		let field;
		await Word.run(async (context) => {
			const selection = context.document.getSelection().getRange();
			field = selection.insertField('Start', 'Addin');
			field.code = `ADDIN test`;
			field.result.insertHtml("<b>Test1</b>", "Start");
			await context.sync();
			field.code = `ADDIN test1after`
			field.track();
			field.result.track();
			tracked.push(field);
			tracked.push(field.result);
			await context.sync();
		});
		await Word.run(tracked, async (context) => {
			// NOTE Uncomment to make this work
			// let fields = context.document.body.fields.load({ result: { text: true } });
			// await context.sync();
			// field = fields.items[0];
			field.result.insertText("", "Replace");
			await context.sync();
			// Throws Sorry, something went wrong. Check the OfficeExtension.Error.debugInfo for more information. 
		});
	}
	catch (e) {
		handleError(e)
	}
	if (event) {
		event.completed();
	}
}

// Fixed
g.testFieldTextReplace = async function(event) {
	await Word.run(async (context) => {
		const selection = context.document.getSelection().getRange();
		const field = selection.insertField('Replace', 'Addin');
		field.code = `ADDIN test`;
		field.result.insertText("Test", "Replace");
		await context.sync();
		field.load({
			result: {
				text: true
			}
		});
		field.result.insertText("Test2", "Replace");
		await context.sync();
	});
	if (event) {
		event.completed();
	}
}

// Fixed
g.testFieldCodePersistAfterTextInsert = async function(event) {
	await Word.run(async (context) => {
		const selection = context.document.getSelection().getRange();
		const field = selection.insertField('Replace', 'Addin');
		field.code = `ADDIN test`;
		result = field.result.insertText("Test", "Replace");
		await context.sync();
		field.code = `ADDIN different`;
		field.result.insertText("Test2", "Replace");
		await context.sync();
	});
	await Word.run(async (context) => {
		const fields = context.document.body.fields;
		fields.load('items');
		await context.sync();
		const field = fields.items[0];
		field.load('code');
		await context.sync();
		console.log(`Field code should be "ADDIN different": ${field.code}`)
	});
	if (event) {
		event.completed();
	}
}
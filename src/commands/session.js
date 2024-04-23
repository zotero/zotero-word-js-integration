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

const FIELD_LOAD_OPTIONS = {
	code: true,
	result: {
		text: true
	}
};
const FIELD_PREFIX = "ADDIN ZOTERO_";
const FIELD_INSERT_CODE = "TEMP";
const FIELD_PLACEHOLDER = "{Updating}";
const PREF_PREFIX = "ZOTERO_PREF";
const PREF_LENGTH = 255;
const BODY_TYPE_TO_NOTE_TYPE = { "Footnote": 1, "Endnote": 2 }
const NOTE_TYPE_TO_BODY_TYPE = ["MainDoc", "Footnote", "Endnote"];
const PLACEHOLDER_LINK_ID_LENGTH = 6;
const IMPORT_LINK_URL = "https://www.zotero.org/";
const IMPORT_DOC_PREFS_PREFIX = "DOCUMENT_PREFERENCES ";
const EXPORT_DOCUMENT_MARKER = "ZOTERO_TRANSFER_DOCUMENT";

/**
 * A class to handle a single button click which initiates an integration
 * transaction
 * @type {Zotero.Session}
 */
Zotero.Session = class {
	constructor(event, command) {
		this.event = event;
		this.command = command;
		this.trackedObjects = [];
		this.fieldsById = {};
		this.orphanFields = [];
		this.insertTextIntoNote = 0;
	}

	/**
	 * Calls the Zotero connector server to initiate an integration transaction
	 * @param command
	 * @returns {Promise<Result>}
	 */
	async execCommand(command) {
		try {
			var request = await Zotero.HTTP.request("POST", ZOTERO_CONFIG.ZOTERO_URL + 'connector/document/execCommand', {
				body: {
					command: command,
					docId: Office.context.document.url
				},
				headers: { "Content-Type": "application/json" },
				timeout: false
			});
			return this.callFunction(JSON.parse(request.response));
		} catch (e) {
			// Usual response for a request in progress
			if (e.status == 503) {
				Zotero.debug(e.message);
				return;
			}
			else if (e.status == 0) {
				return this.displayAlert(`Word could not communicate with Zotero. Please ensure Zotero is running and try again. If this problem persists, see <a target='_blank' href='https://www.zotero.org/support/word_processor_plugin_troubleshooting'>Word Processor Plugin Troubleshooting</a>`)
			}
			Zotero.logError(e);
		}
		finally {
			this.event.completed();
			await this._untrackAll();
		}
	}

	/**
	 * Transacts between word and Zotero connector integration server.
	 * @param result
	 * @returns {Promise<Result>}
	 */
	async respond(result) {
		try {
			var request = await Zotero.HTTP.request("POST", ZOTERO_CONFIG.ZOTERO_URL + 'connector/document/respond', {
				body: result,
				headers: { "Content-Type": "application/json" },
				timeout: false
			});
			return this.callFunction(JSON.parse(request.response));
		} catch (e) {
			Zotero.logError(e);
		}
	}

	/**
	 * Calls one of the integration functions below
	 * @param request
	 * @returns {Promise<Result>}
	 */
	async callFunction(request) {
		var method = request.command.split('.')[1];
		var args = Array.from(request.arguments);
		var docID = args.splice(0, 1);
		var result;
		let wordRunArgs = [];
		if (this.trackedObjects.length) {
			wordRunArgs = [this.trackedObjects];
		}
		
		try {
			await Word.run(...wordRunArgs, async (context) => {
				this.context = context;
				this.document = context.document;
				try {
					result = await this[method].apply(this, args);
				} finally {
					this.context = this.document = null;
				}
			});
		}
		catch (e) {
			Zotero.debug(`Exception in ${request.command}`);
			result = {
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
			Zotero.debug(result);
			debugger;
		}
		
		
		if (method == 'complete') return result;
		return this.respond(result ? JSON.stringify(result) : 'null');
	}

	/**
	 * Adds a tracked object to a list of tracked objects to be freed later.
	 * You need to track objects to be able to access them across Word.run() and context.sync() calls.
	 * @param wordObject
	 * @private
	 */
	_track(wordObject) {
		this.trackedObjects.push(wordObject);
		wordObject.track();
	}

	/**
	 * Untracks all tracked objects. To be called at the end of a transaction
	 * @returns {Promise<void>}
	 * @private
	 */
	async _untrackAll() {
		if (!this.trackedObjects.length) return;
		this.context.trackedObjects.remove(this.trackedObjects);
		await this.trackedObjects[0].context.sync();
	}

	async getDocument() {
		return this.getActiveDocument();
	}

	async getActiveDocument() {
		return {
			documentID: Office.context.document.url,
			primaryFieldType: 'Field',
			secondaryFieldType: 'Bookmark',
			outputFormat: 'html',
			supportedNotes: ['footnotes', 'endnotes'],
			supportsImportExport: true,
			supportsTextInsertion: true,
			supportsCitationMerging: true,
			processorName: "Microsoft Word"
		}
	}

	async getDocumentData() {
		let range = this.document.body.getRange();
		let ranges = range.getTextRanges(["\n"], true);
		ranges.load({ select: "text", top: 1 });
		await this._sync();
		if (ranges.items[0] && ranges.items[0].text.startsWith(EXPORT_DOCUMENT_MARKER)) {
			return EXPORT_DOCUMENT_MARKER;
		}
		const properties = this.document.properties.customProperties;
		properties.load({$all: true});
		await this._sync();
		let pref_pieces = [];
		for (let item of properties.items) {
			if (item.key.startsWith(PREF_PREFIX)) {
				pref_pieces.push(item);
			}
		}
		pref_pieces.sort((a, b) => a.key > b.key);
		return pref_pieces.map(i => i.value).join('');
	}

	async setDocumentData(data) {
		const properties = this.document.properties.customProperties;
		for (let i = 1; data.length; i++) {
			let slice = data.slice(0, PREF_LENGTH)
			properties.add(`${PREF_PREFIX}_${i}`, slice);
			data = data.slice(PREF_LENGTH);
		}
		await this._sync();
	}

	async activate(force) {
		window.focus();
	}

	async cleanup() {}

	async complete() {}

	async displayAlert(text, icon=0, buttons=0) {
		const buttonMapping = [
			["OK"],
			["Ok", "Cancel"],
			["Yes", "No"],
			["Yes", "No", "Cancel"]
		]
		text = encodeURIComponent(text);
		buttons = encodeURIComponent(JSON.stringify(buttonMapping[buttons]));
		const dialogUrl = window.location.origin + `/dialog.html?text=${text}&icon=${icon}&buttons=${buttons}`;
		return new Promise((resolve, reject) => {
			Office.context.ui.displayDialogAsync(dialogUrl, { displayInIframe: true, width: 60, height: 40 }, (asyncResult) => {
				if (asyncResult.error) {
					reject(new Error(`Office.ui.displayDialogAsync error ` + JSON.stringify(asyncResult.error)));
				}
				const dialog = asyncResult.value;
				dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
					dialog.close();
					resolve(arg.message);
				});
				dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
					// Dialog closed by user closing the window
					if (arg.error === 12006) {
						resolve(0)
					}
				})
			});
		})
		
	}

	async getFields() {
		if (this.fields) {
			// This is highly annoying and maybe somewhat bad for performance,
			// but there is NO way to identify a field
			// retrieved from insert/selection and one from a field collection
			// by comparing IDs or something, and the only way to check if things are
			// equal is to use the *ASYNC* range comparison command.
			if (this.orphanFields.length) {
				let comparisons = this.orphanFields.map(_ => []);
				this.orphanFields.forEach((orphanField, idx) => {
					for (let field of this.fields) {
						comparisons[idx].push(field.wordField.result.compareLocationWith(orphanField.wordField.result));
					}
				});
				await this._sync();
				comparisons.forEach((comparison, orphanIdx) => {
					let fieldIdx = comparison.findIndex(c => c.value === "Equal");
					if (fieldIdx === -1) {
						throw new Error ('Orphan Field not found when retrieving all fields');
					}
					let orphanField = this.orphanFields[orphanIdx]
					let field = this.fields[fieldIdx];
					// Use the orphanField id for this one and remap the field returned here to be
					// the "main" one.
					field.id = orphanField.id;
					delete this.fieldsById[field.id];
					this.fieldsById[orphanField.id] = field;
				});
				this.orphanFields = [];
			}
			// If we keep old tracked objects from other calls Word Online breaks
			// with a "something went wrong" prompt, needs a reload and changes are lost.
			if (this.oldTrackedObjects && this.oldTrackedObjects.length) {
				this.context.trackedObjects.remove(this.oldTrackedObjects);
			}
			return this.fields;
		}
		this.oldTrackedObjects = this.trackedObjects;
		this.trackedObjects = [];
		const body = this.document.body;
		let fields = body.fields.getByTypes([Word.FieldType.addin]);
		fields = fields.load(FIELD_LOAD_OPTIONS);
		this._track(fields);
		let footnotes = body.footnotes.load(['items', 'body/type']);
		this._track(footnotes);
		let endnotes = body.endnotes.load(['items', 'body/type']);
		this._track(endnotes);
		await this._sync();
		
		footnotes.items.forEach(note => note.body.fields.load(FIELD_LOAD_OPTIONS));
		endnotes.items.forEach(note => note.body.fields.load(FIELD_LOAD_OPTIONS));
		await this._sync();
		
		let filterNotes = (notes) => {
			let notesHaveZoteroFields = notes.map((note) => {
				let fields = note.body.fields.load(FIELD_LOAD_OPTIONS);
				return fields.items.some(field => field.code.trim().startsWith(FIELD_PREFIX));
			});
			return notes.filter((note, idx) => notesHaveZoteroFields[idx]);
		};
		footnotes = await filterNotes(footnotes.items);
		endnotes = await filterNotes(endnotes.items);
		
		fields = fields.items.filter(field => field.code.trim().startsWith(FIELD_PREFIX));
		fields = await this._sortNotesIntoFields(fields, footnotes)
		fields = await this._sortNotesIntoFields(fields, endnotes)
		
		let getZoteroFieldsFromWordFields = (field, noteType=0, note=null) => {
			if (typeof field.code !== "undefined") {
				if (field.code.trim().startsWith(FIELD_PREFIX)) {
					this._track(field);
					return [this._wordFieldToField(field, noteType, note)];
				}
				else return [];
			}
			let fields = [];
			for (let noteField of field.body.fields.items) {
				this._track(field.body.fields);
				fields = fields.concat(getZoteroFieldsFromWordFields(noteField, BODY_TYPE_TO_NOTE_TYPE[field.body.type], note));
			}
			return fields;
		}
		
		this.fields = [];
		for (let field of fields) {
			this.fields = this.fields.concat(getZoteroFieldsFromWordFields(field))
		}
		
		let adjacency = this.fields.map(_ => ({ value: false }));
		for (let i = 0; i < this.fields.length - 1; i++) {
			let fieldA = this.fields[i];
			let fieldB = this.fields[i+1];
			adjacency[i] = fieldA.wordField.result.compareLocationWith(fieldB.wordField.result);
		}
		await this._sync();
		this.fields.forEach((field, idx) => {
			field.adjacent = adjacency[idx].value === "AdjacentBefore";
		});
		
		return this.getFields();
	}

	async setBibliographyStyle(firstLineIndent, bodyIndent, lineSpacing, entrySpacing,
										 tabStops, tabStopsCount) {
		const styles = this.document.getStyles();
		let style = styles.getByNameOrNullObject(Word.BuiltInStyleName.bibliography);
		await this._sync();
		if (style.isNullObject) {
			// No bibliography style in Word Online!
			style = this.document.addStyle(Word.BuiltInStyleName.bibliography, 'Paragraph');
			await this._sync();
		}
		style.load();
		await this._sync();
		const paragraphFormat = style.paragraphFormat;
		// TODO Word Online/API broken
		// https://github.com/OfficeDev/office-js/issues/3619
		paragraphFormat.firstLineIndent = Math.max(0, firstLineIndent / 20);
		paragraphFormat.leftIndent = bodyIndent / 20;
		paragraphFormat.lineSpacing = lineSpacing / 20;
		paragraphFormat.spaceAfter = entrySpacing / 20;

		// Set tab stops
		// TODO: Missing API reported https://github.com/OfficeDev/office-js/issues/3585
		await this._sync();
	}

	async canInsertField(fieldType) {
		const selection = this.document.getSelection();
		selection.parentBody.load('type');
		await this._sync();
		const type = selection.parentBody.type;
		return (fieldType !== 'Bookmark' && ["Footnote", "Endnote"].includes(type))
			|| type === "MainDoc";
	}
	
	async cursorInField(fieldType) {
		const selection = this.document.getSelection();
		let fields = selection.fields.getByTypes(["Addin"])
		fields = fields.load(FIELD_LOAD_OPTIONS);
		selection.parentBody.load('type');
		await this._sync();
		let noteType = BODY_TYPE_TO_NOTE_TYPE[selection.parentBody.type] || 0;
		this._track(fields);
		for (let field of fields.items) {
			if (field.code.trim().startsWith(FIELD_PREFIX)) {
				this._track(field);
				this._track(field.result);
				await this._sync();
				let note;
				if (noteType) {
					note = await this._getNoteFromBody(field.result.parentBody);
				}
				return this._wordFieldToField(field, noteType, note, true);
			}
		}
		// Unfortunately if the selection is collapsed no fields "in selection" are returned
		// so if a cursor is sitting in a field it's not returned above.
		// We will instead get a range between the start of the body and selection, and
		// another range between the selection and the end of the body and compare the
		// last and first respective fields. If they match, the cursor is in that field.
		const startRange = selection.parentBody.getRange("Start");
		const endRange = selection.parentBody.getRange("End");
		const startToSelectionRange = startRange.expandTo(selection);
		const selectionToEndRange = selection.expandTo(endRange);
		const startFields = startToSelectionRange.fields.getByTypes(["Addin"]);
		const endFields = selectionToEndRange.fields.getByTypes(["Addin"]);
		startFields.load(FIELD_LOAD_OPTIONS);
		endFields.load(FIELD_LOAD_OPTIONS);
		await this._sync();
		const f1 = startFields.items[startFields.items.length-1]
		const f2 = endFields.items[0]
		if (f1 && f2) {
			let comparison = f1.result.compareLocationWith(f2.result);
			await this._sync();
			if (comparison.value === "Equal" && f1.code.trim().startsWith(FIELD_PREFIX)) {
				f1.result.parentBody.load('type');
				this._track(startFields);
				this._track(f1);
				this._track(f1.result);
				await this._sync();
				let note;
				if (noteType) {
					note = await this._getNoteFromBody(f1.result.parentBody);
				}
				return this._wordFieldToField(f1, noteType, note, true);
			}
		}
		return null;
	}


	async insertField(_fieldType, noteType, insertRange) {
		this.noteType = noteType;
		const selection = this.document.getSelection();

		let note;
		if (!insertRange) {
			insertRange = selection;
			selection.parentBody.load('type');
			await this._sync();
			if (noteType && selection.parentBody.type !== NOTE_TYPE_TO_BODY_TYPE[noteType]) {
				if (noteType === 1) {
					note = selection.insertFootnote('');
				}
				else {
					note = selection.insertEndnote('');
				}
				insertRange = note.body.getRange("End");
			}
		}
		
		const field = insertRange.insertField('Replace', 'Addin');
		field.code = `${FIELD_PREFIX}${FIELD_INSERT_CODE}`;
		field.result.insertText(FIELD_PLACEHOLDER, "Replace");
		// TODO: Cannot use FIELD_LOAD_OPTIONS due to a bug
		// https://github.com/OfficeDev/office-js/issues/3615
		field.load(["type", "code"]);
		field.result.load("text");
		field.result.parentBody.load("type");
		this._track(field);
		this._track(field.result);
		await this._sync();
		
		if (noteType && !note) {
			note = await this._getNoteFromBody(field.result.parentBody);
		}	
		return this._wordFieldToField(field, noteType, note, true);
	}

	async insertText(text) {
		const selection = this.document.getSelection();
		let insertRange = selection;
		selection.parentBody.load('type');
		await this._sync();
		// If at the start of operation cursor was in an empty footnote
		// insertField was called, then field.delete, which deleted the
		// footnote. As such we need to reinsert the footnote here.
		if (this.insertTextIntoNote && selection.parentBody.type === "MainDoc") {
			let note;
			if (this.insertTextIntoNote === 1) {
				note = selection.insertFootnote('');
			}
			else {
				note = selection.insertEndnote('');
			}
			insertRange = note.body.getRange("End");
		}
		insertRange.insertHtml(text);
		insertRange.parentBody.getRange().select("End");
	}

	async convertPlaceholdersToFields(placeholderIDs, noteType) {
		const selection = this.document.getSelection();
		let ranges = selection.parentBody.getRange().getHyperlinkRanges();
		ranges.load(['items', 'hyperlink']);
		await this._sync();
		let fields = [];
		for (let range of ranges.items) {
			const id = range.hyperlink.substring(range.hyperlink.length - PLACEHOLDER_LINK_ID_LENGTH);
			if (placeholderIDs.indexOf(id) === -1) continue;
			await this._sync();
			fields.push(await this.insertField("Field", noteType, range));
		}
		this.fields = null;
		return fields;
	}

	async convert(fieldIDs, fieldType, fieldNoteTypes) {
		const fields = await this.getFields();
		if (!fields.length) return;
		if (fields[0].noteType !== fieldNoteTypes[0]) {
			fieldIDs = new Set(fieldIDs);
			if (fieldNoteTypes[0] > 0) {
				await this.inlineToNotes(fieldIDs, fieldNoteTypes);
			}
			else {
				await this.notesToInline(fieldIDs);
			}
		}
		this.fields = null;
	}
	
	async inlineToNotes(fieldIDs, fieldNoteTypes) {
		let insertFunctionNames = {
			1: "insertFootnote",
			2: "insertEndnote"
		}
		let fields = await this.getFields();
		// Reverse sort to not upset doc during update.
		// Ignore fields already in notes.
		fields = fields.reverse().filter(f => !f.noteType && fieldIDs.has(f.id));
		fieldNoteTypes.reverse();
		
		let insertNotes = [];
		fields.forEach((field, index) => {
			const insertNoteType = fieldNoteTypes[index]
			const functionName = insertFunctionNames[insertNoteType];
			const note = field.wordField.result[functionName]();
			note.track();
			insertNotes.push(note);
			field.wordField.delete();
		})
		await this._sync();
		
		fields.forEach((field, index) => {
			const note = insertNotes[index];
			const insertRange = note.body.getRange('End');
			const wordField = insertRange.insertField('Replace', 'Addin');
			wordField.code = `${FIELD_PREFIX}${field.code}`;
			wordField.result.insertText(field.text, "Replace");
			note.untrack();
		});
		await this._sync();
	}
	
	async notesToInline(fieldIDs) {
		let fields = await this.getFields();
		// Reverse sort to not upset doc during update.
		// Ignore fields already inline (shouldn't be any).
		fields = fields.reverse().filter(f => f.noteType && fieldIDs.has(f.id));
		
		let noteRanges = []
		fields.forEach(f => {
			let range = f.wordNote.body.getRange();
			range.load('text');
			noteRanges.push(range);
		});
		await this._sync();
		
		fields.forEach((field, index) => {
			const noteRange = noteRanges[index];
			// See note in delete() for why we're comparing text length and not ranges.
			if (field.wordField.result.text.trim().length !== noteRange.text.trim().length) {
				return;
			}
		
			const wordField = field.wordNote.reference.insertField('End', 'Addin');
			wordField.code = `${FIELD_PREFIX}${field.code}`;
			wordField.result.insertText(field.text, "Replace");
			field.wordNote.delete();
		});
		await this._sync();
	}
	
	async importDocument() {
		let result = false;
		let ranges = [
			this.document.body.getRange(),
			this.document.getFootnoteBody().getRange(),
			this.document.getEndnoteBody().getRange()
		];
		for (let range of ranges) {
			let hyperlinkRanges = range.getHyperlinkRanges();
			hyperlinkRanges.track();
			hyperlinkRanges.load(['items', 'hyperlink', 'text']);
			await this._sync();
			// This is bad code we shouldn't need but the babel transpiler is generating
			// wrong mappings here, so we're stuck with it.
			if (!hyperlinkRanges.items || !hyperlinkRanges.items.length) continue;
			for (let hyperlinkRange of hyperlinkRanges.items.reverse()) {
				if (!hyperlinkRange.hyperlink.startsWith(IMPORT_LINK_URL)) continue;
				const code = hyperlinkRange.text;
				if (code.startsWith(IMPORT_DOC_PREFS_PREFIX)) {
					await this.setDocumentData(code.substring(IMPORT_DOC_PREFS_PREFIX.length));
					result = true;
					hyperlinkRange.insertText("", "Replace");
					continue;
				}
				hyperlinkRange.hyperlink = "";
				let field = hyperlinkRange.insertField("Replace", "Addin");
				field.code = FIELD_PREFIX + code;
				field.result.insertText(FIELD_PLACEHOLDER, "Replace");
			}
			await this._sync();
			hyperlinkRanges.untrack();
		}
		
		if (!result) return result;
		
		// Remove the first 4 paragraphs - import marker, instructions and empty paragraphs.
		let range = this.document.body.getRange();
		ranges = range.getTextRanges(["\n"]);
		ranges.load({ top: 4 });
		await this._sync();
		range = range.getRange('Start');
		range = range.expandTo(ranges.items[3]);
		range.insertText("", "Replace");
		await this._sync();
		
		return result;
	}

	async exportDocument(_, importInstructions) {
		// Document data
		let docData = await this.getDocumentData();
		let insertRange = this.document.body.getRange('End');
		insertRange = insertRange.insertText("\n", "End");
		insertRange = insertRange.insertText(IMPORT_DOC_PREFS_PREFIX + docData, "End");
		insertRange.hyperlink = IMPORT_LINK_URL;
			
		let fields = (await this.getFields()).reverse();
		for (let field of fields) {
			let range = field.wordField.result.getRange('End').insertText(field.code);
			range.hyperlink = IMPORT_LINK_URL;
			field.wordField.delete();
		}
		await this._sync();
	
		// Import instructions
		insertRange = this.document.body.getRange('Start');
		insertRange = insertRange.insertText("\n\n", "Start");
		insertRange = insertRange.insertText(importInstructions, "Start")
		
		// Export marker
		insertRange = insertRange.insertText("\n\n", "Start");
		insertRange.insertText(EXPORT_DOCUMENT_MARKER, "Start");
		await this._sync();
	}	

	async setText(fieldID, text) {
		const field = this.fieldsById[fieldID];
		// TODO: Broken upstream, see https://github.com/OfficeDev/office-js/issues/3613
		// field.wordField.result.insertHtml(text, "Replace");
		field.wordField.result.insertText("", "Replace");
		field.wordField.result.insertHtml(text, "Start");
		if (field.code.startsWith("BIBL")) {
			const style = this.document.getStyles().getByNameOrNullObject(Word.BuiltInStyleName.bibliography);
			await this._sync();
			style.load('builtIn');
			await this._sync();
			if (style.isNullObject) {
				// No bibliography style in Word Online!
				throw new Error("Bibliography style not set before inserting bibliography")
			} else {
				if (style.builtIn) {
					field.wordField.result.styleBuiltIn = Word.BuiltInStyleName.bibliography;
				}
				else {
					field.wordField.result.style = Word.BuiltInStyleName.bibliography;
				}
			}
		}
		await this._sync();
	}

	async setCode(fieldID, code) {
		const field = this.fieldsById[fieldID];
		field.wordField.code = `${FIELD_PREFIX}${code}`;
		field.code = code;
		await this._sync();
	}

	async delete(fieldID) {
		const field = this.fieldsById[fieldID];
		const parentBody = field.wordField.result.parentBody;
		parentBody.load('type');
		await this._sync();
		if (parentBody.type === "MainDoc") {
			field.wordField.result.insertText("", "Replace");
		}
		else {
			field.wordField.result.load('text');
			let noteRange = field.wordNote.body.getRange();
			noteRange.load('text');
			await this._sync();
			// The range.compareWithLocation function won't do us good here, since we cannot construct custom
			// ranges that would account for an empty space.
			// TODO: This does not work correctly right now due to an API bug:
			// https://github.com/OfficeDev/office-js/issues/3591
			if (field.wordField.result.text.trim().length === noteRange.text.trim().length) {
				this.insertTextIntoNote = BODY_TYPE_TO_NOTE_TYPE[field.wordNote.type];
				field.wordNote.delete();
			}
			else {
				field.wordField.result.insertText("", "Replace");
			}
		}
		if (this.fields) {
			this.fields.splice(this.fields.indexOf(field), 1);
		}
		delete this.fieldsById[field.id];
		await this._sync();
	}

	async removeCode(fieldID) {
		const field = this.fieldsById[fieldID];
		field.wordField.delete();
		await this._sync();
	}

	async select(fieldID) {
		const field = this.fieldsById[fieldID];
		field.wordField.result.select();
		await this._sync();
	}

	async _sync() {
		return this.context.sync();
	}
	
	// Comparing ranges in Word JS API is async, so sorting things is quite complicated.
	// Still it will take log(n) async operations to sort two sorted lists into one another
	// which is not the end of the world
	async _sortNotesIntoFields(fields, notes) {
		if (!fields.length) return notes;
		let areSorted = false;
		let compareValues = notes.map(() => ({ value: false }));
		let noteSort = notes.map(() => ({ lower: 0, upper: fields.length }))
		while (true) {
			for (let i = 0; i < notes.length; i++) {
				if (compareValues[i].value !== false) {
					const { lower, upper } = noteSort[i];
					const diff = upper - lower;
					if (compareValues[i].value === "After") {
						noteSort[i].lower = lower + Math.floor((diff)/2.) + (diff === 1 ? 1 : 0);
					} else {
						noteSort[i].upper = lower + Math.floor((diff)/2.) - (diff === 1 ? 1 : 0);
					}
				}
				const { lower, upper } = noteSort[i];
				if (lower === upper) continue;
				const compIdx = lower + Math.floor((upper - lower)/2.)
				if (compIdx >= fields.length) {
					compareValues[i] = { value: "Before" };
					continue;
				}
				const field = fields[compIdx];
				let fieldRange;
				if (typeof field.code != 'undefined') {
					fieldRange = field.result
				} else {
					fieldRange = field.reference;
				}
				compareValues[i] = notes[i].reference.compareLocationWith(fieldRange);
			}
			areSorted = noteSort.every(status => status.lower === status.upper);
			if (areSorted) break;
			await this._sync();
		}
		// Insert in reverse
		notes.reverse();
		noteSort.reverse();
		notes.forEach((note, idx) => {
			fields.splice(noteSort[idx].lower, 0, note)
		});
		return fields;
	}
	
	_wordFieldToField(wordField, noteType, wordNote, orphan=false) {
		let id = randomString();
		const field = {
			code: wordField.code.trim().substr(FIELD_PREFIX.length),
			noteType,
			text: wordField.result.text,
			wordField,
			wordNote,
			id
		}
		if (orphan) {
			this.orphanFields.push(field);
		}
		this.fieldsById[id] = field;
		return field;
	}
	
	async _getNoteFromBody(body) {
		let notes;
		if (body.type === "Footnote") {
			notes = this.document.body.footnotes;
		}
		else {
			notes = this.document.body.endnotes;
		}
		notes.load('items');
		await this._sync();
		let comparisons = [];
		let noteRange = body.getRange();
		for (let note of notes.items) {
			comparisons.push(noteRange.compareLocationWith(note.body.getRange()));
		}
		await this._sync();
		return notes.items.find((_, idx) => {
			return comparisons[idx].value === "Equal";
		});
	}
}

function randomString(len, chars) {
	if (!chars) {
		chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
	}
	if (!len) {
		len = 8;
	}
	var randomstring = '';
	for (var i=0; i<len; i++) {
		var rnum = Math.floor(Math.random() * chars.length);
		randomstring += chars.substring(rnum,rnum+1);
	}
	return randomstring;
}

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
}
const FIELD_PREFIX = "ADDIN ZOTERO_";
const FIELD_INSERT_CODE = "TEMP";
const FIELD_PLACEHOLDER = "{Updating}";
const BODY_TYPE_TO_NOTE_TYPE = { "Footnote": 1, "Endnote": 2 }
const NOTE_TYPE_TO_BODY_TYPE = ["MainDoc", "Footnote", "Endnote"];

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
			else if (e.status == 404) {
				Zotero.confirm({
					title: Zotero.getString('upgradeApp', ZOTERO_CONFIG.CLIENT_NAME),
					message: Zotero.getString(
						'integration_error_clientUpgrade',
						ZOTERO_CONFIG.CLIENT_NAME + ' 5.0.46'
					),
					button2Text: "",
				});
			}
			else if (e.status == 0) {
				var connectorName = Zotero.getString('appConnector', ZOTERO_CONFIG.CLIENT_NAME);
				Zotero.confirm({
					title: Zotero.getString('error_connection_isAppRunning', ZOTERO_CONFIG.CLIENT_NAME),
					message: Zotero.getString(
							'integration_error_connection',
							[connectorName, ZOTERO_CONFIG.CLIENT_NAME]
						)
						+ '<br /><br />'
						+ Zotero.Inject.getConnectionErrorTroubleshootingString(),
					button2Text: "", 
				});
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
			Zotero.logError(e);
			debugger;
			result = {
				error: e.type || `Connector Error`,
				message: e.message,
				stack: e.stack
			}
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
		for (let object of this.trackedObjects) {
			object.untrack();
		}
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
		const properties = this.document.properties.customProperties;
		properties.load({$all: true});
		await this._sync();
		let pref_pieces = [];
		for (let item of properties.items) {
			if (item.key.startsWith(ZOTERO_CONFIG.PREF_PREFIX)) {
				pref_pieces.push(item);
			}
		}
		pref_pieces.sort((a, b) => a.key > b.key);
		return pref_pieces.map(i => i.value).join('');
	}

	async setDocumentData(data) {
		const properties = this.document.properties.customProperties;
		for (let i = 1; data.length; i++) {
			let slice = data.slice(0, ZOTERO_CONFIG.PREF_MAX_LENGTH)
			properties.add(`${ZOTERO_CONFIG.PREF_PREFIX}_${i}`, slice);
			data = data.slice(ZOTERO_CONFIG.PREF_MAX_LENGTH);
		}
		await this._sync();
	}

	async activate(force) {
		window.focus();
	}

	async cleanup() {}

	async complete() {}

	async displayAlert(text, icons, buttons) {
		Zotero.confirm(JSON.stringify({text, icons, buttons}));
		// TODO
		// var result = await Zotero.GoogleDocs.UI.displayAlert(text, icons, buttons);
		// if (buttons < 3) {
		// 	return result % 2;
		// } else {
		// 	return 3 - result;
		// }
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
					this.fields[fieldIdx].id = this.orphanFields[orphanIdx].id;
					delete this.fieldsById[this.orphanFields[orphanIdx].id];
				});
				this.orphanFields = [];
			}
			return this.fields;
		}
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
		
		let getZoteroFieldsFromWordFields = (field, noteType=0) => {
			if (typeof field.code !== "undefined") {
				if (field.code.trim().startsWith(FIELD_PREFIX)) {
					this._track(field);
					this._track(field.result);
					return [this._wordFieldToField(field, noteType)];
				}
				else return [];
			}
			let fields = [];
			for (let noteField of field.body.fields.items) {
				fields = fields.concat(getZoteroFieldsFromWordFields(noteField, BODY_TYPE_TO_NOTE_TYPE[field.body.type]));
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
		// TODO: Always returns "Equals". Reported https://github.com/OfficeDev/office-js/issues/3584 
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
				return this._wordFieldToField(field, noteType, true);
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
				this._track(startFields);
				this._track(f1);
				this._track(f1.result);
				await this._sync();
				return this._wordFieldToField(f1, noteType, true);
			}
		}
		return null;
	}


	async insertField(fieldType, noteType) {
		const selection = this.document.getSelection();
		
		let insertRange = selection;
		let note;
		selection.parentBody.load('type');
		if (noteType && selection.parentBody.type !== NOTE_TYPE_TO_BODY_TYPE[noteType]) {
			if (noteType === 1) {
				note = selection.insertFootnote('');
			}
			else {
				note = selection.insertEndnote('');
			}
			insertRange = note.body.getRange();
		}
		
		const field = insertRange.insertField('Replace', 'Addin');
		field.code = `${FIELD_PREFIX}${FIELD_INSERT_CODE}`;
		field.result.insertText(FIELD_PLACEHOLDER, "Replace");
		// TODO: Cannot use FIELD_LOAD_OPTIONS due to a bug
		// https://github.com/OfficeDev/office-js/issues/3615
		field.load(["type", "code"]);
		field.result.load("text");
		this._track(field);
		this._track(field.result);
		await this._sync();
		
		return this._wordFieldToField(field, noteType, true);
	}

	async insertText(text) {
		// TODO
	}

	async convertPlaceholdersToFields(placeholderIDs, noteType) {
		// TODO
	}

	async convert(fieldIDs, fieldType, fieldNoteTypes) {
		// TODO
	}
	
	async importDocument() {
		// TODO
	}

	async exportDocument() {
		// TODO
	}	

	async setText(fieldID, text) {
		const field = this.fieldsById[fieldID];
		// TODO: Broken upstream, see https://github.com/OfficeDev/office-js/issues/3613
		// field.wordField.result.insertHtml(text, "Replace");
		field.wordField.result.insertText(text.replace(/\n/g, ""), "Replace");
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
		field.wordField.result.insertText("", "Replace");
		this.fields.remove(field);
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
		notes.sort(() => -1);
		noteSort.sort(() => -1);
		notes.forEach((note, idx) => {
			fields.splice(noteSort[idx].lower, 0, note)
		});
		return fields;
	}
	
	_wordFieldToField(wordField, noteType, orphan=false) {
		let id = randomString();
		const field = {
			code: wordField.code.trim().substr(FIELD_PREFIX.length),
			noteType,
			wordField: wordField,
			text: wordField.result.text,
			id
		}
		if (orphan) {
			this.orphanFields.push(field);
		}
		this.fieldsById[id] = field;
		return field;
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

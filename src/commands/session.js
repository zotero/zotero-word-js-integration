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
	type: true,
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
const ID_REGEXP = /{[a-zA-Z0-9\-]+}/;

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
		await this.sync();
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
		await this.sync();
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
		if (this.fields) return this.fields;
		const body = this.document.body;
		let fields = body.fields.getByTypes([Word.FieldType.addin]);
		fields = fields.load(FIELD_LOAD_OPTIONS);
		this._track(fields);
		let footnotes = body.footnotes.load('items/type');
		this._track(footnotes);
		let endnotes = body.endnotes.load('items/type');
		this._track(endnotes);
		await this.sync();
		
		footnotes.items.forEach(note => note.body.fields.load(FIELD_LOAD_OPTIONS));
		endnotes.items.forEach(note => note.body.fields.load(FIELD_LOAD_OPTIONS));
		await this.sync();
		
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
					return [this._wordFieldToField(field, noteType)];
				}
				else return [];
			}
			let fields = [];
			for (let noteField of field.body.fields.items) {
				fields = fields.concat(getZoteroFieldsFromWordFields(noteField, BODY_TYPE_TO_NOTE_TYPE[noteField.type]));
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
		await this.sync();
		// TODO: Always returns "Equals". Reported https://github.com/OfficeDev/office-js/issues/3584 
		this.fields.forEach((field, idx) => {
			field.adjacent = adjacency[idx].value === "AdjacentBefore";
		});
		
		return this.fields;
	}

	setBibliographyStyle(firstLineIndent, bodyIndent, lineSpacing, entrySpacing,
										 tabStops, tabStopsCount) {
		const styles = this.document.getStyles();
		const style = styles.getByName(Word.BuiltInStyleName.bibliography);
		style.load({
			paragraphFormat: { $all: true }
		});
		const paragraphFormat = style.paragraphFormat;
		paragraphFormat.firstLineIndent =  firstLineIndent / 20.0;
		paragraphFormat.leftIndent = bodyIndent / 20.0;
		paragraphFormat.lineSpacing = lineSpacing / 20.0;
		paragraphFormat.spaceAfter = entrySpacing / 20.0;
		
		// Set tab stops
		// TODO: Missing API reported https://github.com/OfficeDev/office-js/issues/3585
	}

	async canInsertField(fieldType) {
		const selection = this.document.getSelection();
		selection.parentBody.load('type');
		await this.sync();
		const type = selection.parentBody.type;
		return (fieldType !== 'Bookmark' && ["Footnote", "Endnote"].includes(type))
			|| type === "MainDoc";
	}
	
	async cursorInField(fieldType) {
		const selection = this.document.getSelection();
		let fields = selection.fields.getByTypes(["Addin"])
		fields = fields.load(FIELD_LOAD_OPTIONS);
		selection.parentBody.load('type');
		await this.sync();
		let noteType = BODY_TYPE_TO_NOTE_TYPE[selection.parentBody.type] || 0;
		for (let field of fields.items) {
			if (field.code.trim().startsWith(FIELD_PREFIX)) {
				this._track(fields);
				this._track(field);
				await this.sync();
				return this._wordFieldToField(field, noteType);
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
		await this.sync();
		const f1 = startFields.items[startFields.items.length-1]
		const f2 = endFields.items[0]
		if (this._getWordObjectID(f1) === this._getWordObjectID(f2)) {
			if (f1.code.trim().startsWith(FIELD_PREFIX)) {
				this._track(startFields);
				this._track(f1);
				await this.sync();
				return this._wordFieldToField(f1, noteType);
			}
		}
		return null;
	}


	async insertField(fieldType, noteType) {
		const selection = this.document.getSelection();
		selection.parentBody.load('type');
		let insertRange = selection;
		if (noteType && selection.parentBody.type !== NOTE_TYPE_TO_BODY_TYPE[noteType]) {
			let note;
			if (noteType === 1) {
				note = selection.insertFootnote('');
			}
			else {
				note = selection.insertEndnote('');
			}
			insertRange = note.body.getRange();
		}
		const field = insertRange.insertField('Replace', 'Addin');
		field.code = `ADDIN ${FIELD_PREFIX} ${FIELD_INSERT_CODE}`;
		field.result.insertText(FIELD_PLACEHOLDER, "Replace");
		field.load(FIELD_LOAD_OPTIONS);
		this._track(field);
		await this.sync();
		return this._wordFieldToField(field, noteType);
	}

	async insertText(text) {
		this.insertingNote = true;
		await Zotero.GoogleDocs.UI.writeText(text);
		await Zotero.GoogleDocs.UI.waitToSaveInsertion();
	}

	async convertPlaceholdersToFields(placeholderIDs, noteType) {
		let document = new Zotero.GoogleDocs.Document(await Zotero.GoogleDocs_API.getDocument(this.documentID));
		let links = document.getLinks();

		let placeholders = [];
		for (let link of links) {
			if (link.url.startsWith(Zotero.GoogleDocs.config.fieldURL) ||
				!link.url.startsWith(Zotero.GoogleDocs.config.noteInsertionPlaceholderURL)) continue;
			let id = link.url.substr(Zotero.GoogleDocs.config.noteInsertionPlaceholderURL.length);
			let index = placeholderIDs.indexOf(id);
			if (index == -1) continue;
			link.id = id;
			link.index = index;
			link.code = "TEMP";
			placeholders.push(link);
		}
		// Sanity check
		if (placeholders.length != placeholderIDs.length){
			throw new Error(`convertPlaceholdersToFields: number of placeholders (${placeholders.length}) do not match the number of provided placeholder IDs (${placeholderIDs.length})`);
		}
		let requestBody = { writeControl: { targetRevisionId: document.revisionId } };
		let requests = [];
		// Sort for update by reverse order of appearance to correctly update the doc
		placeholders.sort((a, b) => b.endIndex - a.endIndex);
		if (noteType == 1 && !placeholders[0].footnoteId) {
			// Insert footnotes (and remove placeholders) (using the Google Docs API we can do that properly!)
			for (let placeholder of placeholders) {
				requests.push({
					createFootnote: {
						location: {
							index: placeholder.startIndex,
						}
					}
				});
				requests.push({
					deleteContentRange: {
						range: {
							startIndex: placeholder.startIndex+1,
							endIndex: placeholder.endIndex+1,
						}
					}
				});
			}
			requestBody.requests = requests;
			let response = await Zotero.GoogleDocs_API.batchUpdateDocument(this.documentID, requestBody);

			// Reinsert placeholders in the inserted footnotes
			requestBody = {};
			requests = [];
			placeholders.forEach((placeholder, index) => {
				// Every second response is from createFootnote
				let footnoteId = response.replies[index * 2].createFootnote.footnoteId;
				requests.push({
					insertText: {
						text: placeholder.text,
						location: {
							index: 1,
							segmentId: footnoteId
						}
					}
				});
				requests.push({
					updateTextStyle: {
						textStyle: {
							link: {
								url: Zotero.GoogleDocs.config.fieldURL + placeholder.id
							}
						},
						fields: 'link',
						range: {
							startIndex: 1,
							endIndex: placeholder.text.length+1,
							segmentId: footnoteId
						}
					}
				});
			});
			requestBody.requests = requests;
			await Zotero.GoogleDocs_API.batchUpdateDocument(this.documentID, requestBody);
		} else {
			for (let placeholder of placeholders) {
				requests.push({
					updateTextStyle: {
						textStyle: {
							link: {
								url: Zotero.GoogleDocs.config.fieldURL + placeholder.id
							}
						},
						fields: 'link',
						range: {
							startIndex: placeholder.startIndex,
							endIndex: placeholder.endIndex,
							segmentId: placeholder.footnoteId
						}
					}
				});
				if (placeholder.text[0] == ' ') {
					requests.push({
						updateTextStyle: {
							textStyle: {},
							fields: 'link',
							range: {
								startIndex: placeholder.startIndex,
								endIndex: placeholder.startIndex+1,
								segmentId: placeholder.footnoteId
							}
						}
					});
				}
			}
			requestBody.requests = requests;
			await Zotero.GoogleDocs_API.batchUpdateDocument(this.documentID, requestBody);
		}
		// Reverse to sort in order of appearance, to make sure getFields returns inserted fields
		// in the correct order 
		placeholders.reverse();
		// Queue insert calls to apps script, where the insertion of field text and code will be finalized
		placeholders.forEach(placeholder => {
			var field = {
				text: placeholder.text,
				code: placeholder.code,
				id: placeholder.id,
				noteIndex: noteType ? this.insertNoteIndex++ : 0
			};
			this.queued.insert.push(field);
		});
		// Returning inserted fields in the order of appearance of placeholder IDs
		return Array.from(this.queued.insert).sort((a, b) => placeholderIDs.indexOf(a.id) - placeholderIDs.indexOf(b.id));
	}

	async convert(fieldIDs, fieldType, fieldNoteTypes) {
		var fields = await this.getFields();
		var fieldMap = {};
		for (let field of fields) {
			fieldMap[field.id] = field;
		}

		this.queued.conversion = true;
		if (fieldMap[fieldIDs[0]].noteIndex != fieldNoteTypes[0]) {
			// Note/intext conversions
			if (fieldNoteTypes[0] > 0) {
				fieldIDs = new Set(fieldIDs);
				let document = new Zotero.GoogleDocs.Document(await Zotero.GoogleDocs_API.getDocument(this.documentID));
				let links = document.getLinks()
					.filter((link) => {
						if (!link.url.startsWith(Zotero.GoogleDocs.config.fieldURL)) return false;
						let id = link.url.substr(Zotero.GoogleDocs.config.fieldURL.length);
						return fieldIDs.has(id) && !link.footnoteId;

					})
					// Sort for update by reverse order of appearance to correctly update the doc
					.reverse();
				let requestBody = { writeControl: { targetRevisionId: document.revisionId } };
				let requests = [];

				// Insert footnotes (and remove placeholders)
				for (let link of links) {
					requests.push({
						createFootnote: {
							location: {
								index: link.endIndex,
							}
						}
					});
					requests.push({
						deleteContentRange: {
							range: {
								startIndex: link.startIndex,
								endIndex: link.endIndex,
							}
						}
					});
				}
				requestBody.requests = requests;
				let response = await Zotero.GoogleDocs_API.batchUpdateDocument(this.documentID, requestBody);

				// Reinsert placeholders in the inserted footnotes
				requestBody = {};
				requests = [];
				links.forEach((link, index) => {
					// Every second response is from createFootnote
					let footnoteId = response.replies[index * 2].createFootnote.footnoteId;
					requests.push({
						insertText: {
							text: link.text,
							location: {
								index: 1,
								segmentId: footnoteId
							}
						}
					});
					requests.push({
						updateTextStyle: {
							textStyle: {
								link: {
									url: link.url
								}
							},
							fields: 'link',
							range: {
								startIndex: 1,
								endIndex: link.text.length+1,
								segmentId: footnoteId
							}
						}
					});
				});
				requestBody.requests = requests;
				await Zotero.GoogleDocs_API.batchUpdateDocument(this.documentID, requestBody);
			} else {
				// To in-text conversions client-side are impossible, because there is no obvious way
				// to make the cursor jump from the footnote section to its corresponding footnote.
				// Luckily, this can be done in Apps Script.
				return Zotero.GoogleDocs_API.run(this.documentID, 'footnotesToInline', [
					fieldIDs,
				]);
			}
		}
	}

	async setText(fieldID, text, isRich) {
		if (!(fieldID in this.queued.fields)) {
			this.queued.fields[fieldID] = {id: fieldID};
		}
		// Fixing Google bugs. Google Docs XML parser ignores spaces between tags
		// e.g. <i>Journal</i> <b>2016</b>.
		// The space above is ignored, so we move it into the previous tag
		this.queued.fields[fieldID].text = text.replace(/(<\s*\/[^>]+>) +</g, ' $1<');
		this.queued.fields[fieldID].isRich = isRich;
	}

	async setCode(fieldID, code) {
		if (!(fieldID in this.queued.fields)) {
			this.queued.fields[fieldID] = {id: fieldID};
		}
		// The speed of updates is highly dependend on the size of
		// field codes. There are a few citation styles that require
		// the abstract field, but they are not many and the speed
		// improvement is worth the sacrifice. The users who need to
		// use the styles that require the abstract field will have to
		// cite items from a common group library.
		var startJSON = code.indexOf('{');
		var endJSON = code.lastIndexOf('}');
		if (startJSON != -1 && endJSON != -1) {
			var json = JSON.parse(code.substring(startJSON, endJSON+1));
			delete json.schema;
			if (json.citationItems) {
				for (let i = 0; i < json.citationItems.length; i++) {
					delete json.citationItems[i].itemData.abstract;
				}
				code = code.substring(0, startJSON) + JSON.stringify(json) + code.substring(endJSON+1);
			}
		}
		this.queued.fields[fieldID].code = code;
	}

	async delete(fieldID) {
		if (this.queued.insert[0] && this.queued.insert[0].id == fieldID) {
			let [field] = this.queued.insert.splice(0, 1);
			await Zotero.GoogleDocs.UI.undo();
			if (field.noteIndex > 0) {
				await Zotero.GoogleDocs.UI.undo();
				await Zotero.GoogleDocs.UI.undo();
				await Zotero.GoogleDocs.UI.undo();
			}
			delete this.queued.fields[fieldID];
			return;
		}
		if (!(fieldID in this.queued.fields)) {
			this.queued.fields[fieldID] = {id: fieldID};
		}
		this.queued.fields[fieldID].delete = true;
	}

	async removeCode(fieldID) {
		if (this.queued.insert && this.queued.insert.id == fieldID) {
			this.queued.insert.removeCode = true;
		}
		if (!(fieldID in this.queued.fields)) {
			this.queued.fields[fieldID] = {id: fieldID};
		}
		this.queued.fields[fieldID].removeCode = true;
		// This call is a part of Unlink Citations, which means that
		// after this there will be no more Zotero links in the file
		Zotero.GoogleDocs.hasZoteroCitations = false;
	}

	async select(fieldID) {
		let fields = await this.getFields();
		let field = fields.find(f => f.id == fieldID);

		if (!field) {
			throw new Error(`Attempting to select field ${fieldID} that does not exist in the document`);
		}
		let url = Zotero.GoogleDocs.config.fieldURL+field.id;
		if (!await Zotero.GoogleDocs.UI.selectText(field.text, url)) {
			Zotero.debug(`Failed to select ${field.text} with url ${url}`);
		}
	}

	async importDocument() {
		delete this.fields;
		return Zotero.GoogleDocs_API.run(this.documentID, 'importDocument');
		Zotero.GoogleDocs.downloadInterceptBlocked = false;
	}

	async exportDocument() {
		await Zotero.GoogleDocs_API.run(this.documentID, 'exportDocument', Array.from(arguments));
		var i = 0;
		Zotero.debug(`GDocs: Clearing fields ${i++}`);
		while (!(await Zotero.GoogleDocs_API.run(this.documentID, 'clearAllFields'))) {
			Zotero.debug(`GDocs: Clearing fields ${i++}`)
		}
		Zotero.GoogleDocs.downloadInterceptBlocked = true;
	}	

	async sync() {
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
			await this.sync();
		}
		// Insert in reverse
		notes.sort(() => -1);
		noteSort.sort(() => -1);
		notes.forEach((note, idx) => {
			fields.splice(noteSort[idx].lower, 0, note)
		});
		return fields;
	}

	/**
	 * This is using an undocumented API and may break. We cannot implement field.equals() without
	 * something like this.
	 * @param object A word object we want an unique tracking/comparing ID for
	 */
	_getWordObjectID(object) {
		try {
			return object._objectPath.objectPathInfo.Id
		}
		catch (e) {
			throw new Error('Failed to retrieve the field id.\n' + e.message);
		}
	}
	
	_wordFieldToField(wordField, noteType) {
		// Tapping into undocumented APIs here, so this might fail
		let id = this._getWordObjectID(wordField);
		return {
			code: wordField.code.trim().substr(FIELD_PREFIX.length),
			noteType,
			wordField: wordField,
			text: wordField.result.text,
			id
		}
	}
}
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

/**
 * A class to handle a single button click which initiates an integration
 * transaction
 * @type {Zotero.Session}
 */
Zotero.Session = class {
	constructor(event, command) {
		this.event = event;
		this.command = command;
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
			await this._releaseContext();
			this.event.completed();
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
		
		if (!this.context) {
			await this._getContext();
		}
		try {
			result = await this[method].apply(this, args);
		} catch (e) {
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
	
	async _getContext() {
		if (this.context) return this.context;
		this._releaseContextDeferred = Zotero.Promise.defer();
		return new Promise((resolve) => {
			Word.run(context => {
				this.context = context;
				resolve(context);
				return this._releaseContextDeferred.promise;
			});
		});
	}
	
	async _releaseContext() {
		if (!this._releaseContextDeferred) return;
		this._releaseContextDeferred.resolve();
		this.context = null;
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
		const properties = this.context.document.properties.customProperties;
		properties.load({$all: true});
		await this.context.sync();
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
		const properties = this.context.document.properties.customProperties;
		for (let i = 1; data.length; i++) {
			let slice = data.slice(0, ZOTERO_CONFIG.PREF_MAX_LENGTH)
			properties.add(`${ZOTERO_CONFIG.PREF_PREFIX}_${i}`, slice);
			data = data.slice(ZOTERO_CONFIG.PREF_MAX_LENGTH);
		}
		await this.context.sync();
	}

	async activate(force) {
		window.focus();
	}

	async cleanup() {}

	async complete() {}

	async displayAlert(text, icons, buttons) {
		Zotero.confirm({text, icons, buttons});
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
		const body = this.context.document.body;
		let fields = body.fields.getByTypes([Word.FieldType.addin,
			Word.FieldType.noteRef])
		fields = fields.load(FIELD_LOAD_OPTIONS);
		await this.context.sync();
		
		// Load note info
		fields.items.forEach(field => {
			if (field.type === "NoteRef") {
				field.result.footnotes.load("items")
				field.result.endnotes.load("items")
			}
		})
		await this.context.sync();
		
		let getZoteroFieldsFromWordFields = (field, noteType=0) => {
			if (field.type === "Addin" && field.code.trim().startsWith(FIELD_PREFIX)) {
				field.track();
				return [{
					code: field.code.trim(),
					noteType,
					wordField: field,
					text: field.result.text
				}];
			}
			let fields = [];
			let note = this._getNoteFromNoteField(field);
			if (!note) return [];
			for (let field of note.body.fields.items) {
				fields = fields.concat(getZoteroFieldsFromWordFields(field, field.type === 'Footnote' ? 1 : 2));
			}
			return fields;
		}
		
		this.fields = [];
		for (let field of fields.items) {
			this.fields = this.fields.concat(getZoteroFieldsFromWordFields(field))
		}
		
		let adjacency = this.fields.map(_ => ({ value: false }));
		for (let i = 0; i < this.fields.length - 1; i++) {
			let fieldA = this.fields[i];
			let fieldB = this.fields[i+1];
			adjacency[i] = fieldA.wordField.result.compareLocationWith(fieldB.wordField.result);
		}
		await this.context.sync();
		// TODO: Always returns "Equals". Reported https://github.com/OfficeDev/office-js/issues/3584 
		this.fields.forEach((field, idx) => {
			field.adjacent = adjacency[idx].value === "AdjacentBefore";
		});
		
		return this.fields;
	}

	setBibliographyStyle(firstLineIndent, bodyIndent, lineSpacing, entrySpacing,
										 tabStops, tabStopsCount) {
		const styles = this.context.document.getStyles();
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

	async insertField(fieldType, noteType) {
		var id = Zotero.Utilities.randomString(Zotero.GoogleDocs.config.fieldKeyLength);
		var field = {
			text: Zotero.GoogleDocs.config.citationPlaceholder,
			code: '{}',
			id,
			noteIndex: noteType ? this.insertNoteIndex : 0
		};

		this.queued.insert.push(field);
		await this._insertField(field, false);
		return field;
	}

	async insertText(text) {
		this.insertingNote = true;
		await Zotero.GoogleDocs.UI.writeText(text);
		await Zotero.GoogleDocs.UI.waitToSaveInsertion();
	}

	/**
	 * Insert a front-side link at selection with field ID in the url. The text and field code
	 * should later be saved from the server-side AppsScript code.
	 *
	 * @param {Object} field
	 */
	async _insertField(field, waitForSave=true, ignoreNote=false) {
		var url = Zotero.GoogleDocs.config.fieldURL + field.id;

		if (field.noteIndex > 0) {
			await Zotero.GoogleDocs.UI.insertFootnote();
		}
		await Zotero.GoogleDocs.UI.insertLink(field.text, url);

		if (!waitForSave) {
			return;
		}
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

	async cursorInField(showOrphanedCitationAlert=false) {
		if (!(this.currentFieldID)) return false;
		this.isInOrphanedField = false;

		var fields = await this.getFields();
		// The call to getFields() might change the selectedFieldID if there are duplicates
		let selectedFieldID = this.currentFieldID = await Zotero.GoogleDocs.UI.getSelectedFieldID();
		for (let field of fields) {
			if (field.id == selectedFieldID) {
				return field;
			}
		}
		if (selectedFieldID && selectedFieldID.startsWith("broken=")) {
			this.isInOrphanedField = true;
			if (showOrphanedCitationAlert === true && !this.orphanedCitationAlertShown) {
				let result = await Zotero.GoogleDocs.UI.displayOrphanedCitationAlert();
				if (!result) {
					throw new Error('Handled Error');
				}
				this.orphanedCitationAlertShown = true;
			}
			return false;
		}
		throw new Error(`Selected field ${selectedFieldID} not returned from Docs backend`);
	}

	async canInsertField() {
		return this.isInOrphanedField || !this.isInLink;
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
	
	_getNoteFromNoteField(field) {
		if (field.result.footnotes.items.length) {
			return field.results.footnotes.items[0];
		} else if (field.results.endnotes.items.length) {
			return field.results.endnotes.items[0];
		}
		return null;
	}
}
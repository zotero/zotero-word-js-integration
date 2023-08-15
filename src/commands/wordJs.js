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
	'setDocPrefs',
	'refresh',
	'unlink'
]
for (let button of buttons) {
	g[button] = generateButtonHandler(button);
}

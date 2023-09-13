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
import {
    provideFluentDesignSystem,
    fluentButton
} from "@fluentui/web-components";

provideFluentDesignSystem()
    .register(
        fluentButton()
    );
    
Office.onReady(function() {
	const params = new URLSearchParams(document.location.search);
	initDialog(params.get('text'), params.get('icon'), JSON.parse(params.get('buttons')))
	
	document.addEventListener('keydown', (event) => {
		if (event.key === "Escape") {
			Office.context.ui.messageParent(0);
		}
	})
});
    
function initDialog(text, _icon, buttons) {
	text = text.replace(/\n/g, '<br/>');
	document.querySelector('#text').innerHTML = text;

	const buttonsElem = document.querySelector('#buttons');
	
	buttons.forEach((button, idx) => {
		if (idx === 2) {
			// Add a separator to display the last button on the left
			const div = document.createElement('div');
			div.style.flex = "1";
			buttonsElem.append(div);
		}
		const buttonElem = document.createElement('fluent-button');
		buttonElem.innerText = button;
		buttonElem.addEventListener('click', function() {
			clickButton(buttons.length - 1 - idx);
		});
		buttonsElem.append(buttonElem);
		if (idx === 0) {
			buttonElem.setAttribute('appearance', "accent");
			buttonElem.focus();
		}
	});
}

function clickButton(buttonIdx) {
	Office.context.ui.messageParent(buttonIdx);
}
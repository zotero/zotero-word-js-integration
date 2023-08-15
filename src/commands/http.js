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


/**
 * Functions for performing HTTP requests, both via XMLHTTPRequest and using a hidden browser
 * @namespace
 */
Zotero.HTTP = new function() {
	this.StatusError = function(xmlhttp, url) {
		this.message = `HTTP request to ${url} rejected with status ${xmlhttp.status}`;
		this.status = xmlhttp.status;
		try {
			this.responseText = typeof xmlhttp.responseText == 'string' ? xmlhttp.responseText : undefined;
		} catch (e) {}
	};
	this.StatusError.prototype = Object.create(Error.prototype);

	this.TimeoutError = function(ms) {
		this.message = `HTTP request has timed out after ${ms}ms`;
	};
	this.TimeoutError.prototype = Object.create(Error.prototype);
	
	/**
	 * Get a promise for a HTTP request
	 *
	 * @param {String} method The method of the request ("GET", "POST", "HEAD", or "OPTIONS")
	 * @param {String}	url				URL to request
	 * @param {Object} [options] Options for HTTP request:<ul>
	 *         <li>body - The body of a POST request</li>
	 *         <li>headers - Object of HTTP headers to send with the request</li>
	 *         <li>debug - Log response text and status code</li>
	 *         <li>logBodyLength - Length of request body to log</li>
	 *         <li>timeout - Request timeout specified in milliseconds [default 15000]</li>
	 *         <li>responseType - The response type of the request from the XHR spec</li>
	 *         <li>responseCharset - The charset the response should be interpreted as</li>
	 *         <li>successCodes - HTTP status codes that are considered successful, or FALSE to allow all</li>
	 *     </ul>
	 * @return {Promise<XMLHttpRequest>} A promise resolved with the XMLHttpRequest object if the
	 *     request succeeds, or rejected if the browser is offline or a non-2XX status response
	 *     code is received (or a code not in options.successCodes if provided).
	 */
	this.request = function(method, url, options = {}) {
		// Default options
		options = Object.assign({
			body: null,
			headers: {},
			debug: false,
			logBodyLength: 1024,
			timeout: 15000,
			responseType: '',
			responseCharset: null,
			successCodes: null
		}, options);
		
		
		let logBody = '';
		if (['GET', 'HEAD'].includes(method)) {
			if (options.body != null) {
				throw new Error(`HTTP ${method} cannot have a request body (${options.body})`)
			}
		} else if(options.body) {
			if (!options.headers) options.headers = {};
			if (typeof options.body == 'object') {
				options.body = JSON.stringify(options.body)
				if (!options.headers["Content-Type"]) {
					options.headers["Content-Type"] = 'application/json'
				}
			}
			
			if (!options.headers["Content-Type"]) {
				options.headers["Content-Type"] = "application/x-www-form-urlencoded";
			}
			else if (options.headers["Content-Type"] == 'multipart/form-data') {
				// Allow XHR to set Content-Type with boundary for multipart/form-data
				delete options.headers["Content-Type"];
			}
					
			logBody = `: ${options.body.substr(0, options.logBodyLength)}` +
					options.body.length > options.logBodyLength ? '...' : '';
			// TODO: make sure below does its job in every API call instance
			// Don't display password or session id in console
			logBody = logBody.replace(/password":"[^"]+/, 'password":"********');
			logBody = logBody.replace(/password=[^&]+/, 'password=********');
		}
		Zotero.debug(`HTTP ${method} ${url}${logBody}`);
		
		var xmlhttp = new XMLHttpRequest();
		xmlhttp.timeout = options.timeout;
		var promise = Zotero.HTTP._attachHandlers(url, xmlhttp, options);
		
		xmlhttp.open(method, url, true);

		for (let header in options.headers) {
			xmlhttp.setRequestHeader(header, options.headers[header]);
		}
		
		xmlhttp.responseType = options.responseType || '';
		
		// Maybe should provide "mimeType" option instead. This is xpcom legacy, where responseCharset
		// could be controlled manually
		if (options.responseCharset) {
			xmlhttp.overrideMimeType("text/plain; charset=" + options.responseCharset);
		}
		
		xmlhttp.send(options.body);
		
		return promise.then(function(xmlhttp) {
			if (options.debug) {
				if (xmlhttp.responseType == '' || xmlhttp.responseType == 'text') {
					Zotero.debug(`HTTP ${xmlhttp.status} response: ${xmlhttp.responseText}`);
				}
				else {
					Zotero.debug(`HTTP ${xmlhttp.status} response`);
				}
			}	
			
			let invalidDefaultStatus = options.successCodes === null && !xmlhttp.responseURL.startsWith("file://") &&
				(xmlhttp.status < 200 || xmlhttp.status >= 300);
			let invalidStatus = Array.isArray(options.successCodes) && !options.successCodes.includes(xmlhttp.status);
			if (invalidDefaultStatus || invalidStatus) {
				throw new Zotero.HTTP.StatusError(xmlhttp, url);
			}
			return xmlhttp;
		});
	}
		
	/**
	 * Adds request handlers to the XMLHttpRequest and returns a promise that resolves when
	 * the request is complete. xmlhttp.send() still needs to be called, this just attaches the
	 * handler
	 *
	 * See {@link Zotero.HTTP.request} for parameters
	 * @private
	 */
	this._attachHandlers = function(url, xmlhttp, options) {
		var deferred = Zotero.Promise.defer();
		xmlhttp.onload = () => deferred.resolve(xmlhttp);
		xmlhttp.onerror = xmlhttp.onabort = function() {
			var e = new Zotero.HTTP.StatusError(xmlhttp, url);
			if (options.successCodes === false) {
				deferred.resolve(xmlhttp);
			} else {
				deferred.reject(e);
			}
		};
		xmlhttp.ontimeout = function() {
			var e = new Zotero.HTTP.TimeoutError(xmlhttp.timeout);
			Zotero.logError(e);
			deferred.reject(e);
		};
		return deferred.promise;
	};
}
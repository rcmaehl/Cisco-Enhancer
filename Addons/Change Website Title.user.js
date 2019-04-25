// ==UserScript==
// @name        Change Website Title
// @description Prepends Website Title based on URL
// @match       *://*.discordapp.com/*
// @match       *://*.facebook.com/*
// @match       *://*.reddit.com/*
// @version     1.4
// @grant       none
// @run-at      document-idle
// ==/UserScript==

var target = document.querySelector('title');
var config = { attributes: true, childList: true, characterData: true }
var observer = new MutationObserver(function(mutations) {
    mutations.forEach(function(mutation) {
        console.log(mutation.type);
      	observer.disconnect();
        if (location.hostname.endsWith("discordapp.com")) {
            document.title="Discord: " + document.title
            observer.observe(target, config);
            console.log("Title Changed");
        } else if (location.hostname.endsWith("facebook.com")) {
            document.title="Facebook: " + document.title
            observer.observe(target, config);
            console.log("Title Changed");
        } else if (location.hostname.endsWith("reddit.com")) {
            document.title="Reddit: " + document.title
            observer.observe(target, config);
            console.log("Title Changed");
        };
        console.log("Path: " + window.location);
    });
});
observer.observe(target, config);

(function() {
    'use strict';
    document.title = "" + document.title
})();
// ==UserScript==
// @name         Auto Presenza Teleskill + Sveglia + Log
// @namespace    http://tampermonkey.net/
// @version      1.4
// @description  Clic automatico su "PROSEGUI" durante la verifica presenza su Teleskill, con sveglia e log
// @match        *://tlive.teleskill.it/*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const parolaTarget = "prosegui";
    const durataSuono = 5000; // durata suono in millisecondi
    const volumeSuono = 0.7;

    let audioInterval = null;

    function normalizeText(text) {
        return text ? text.toLowerCase().trim() : "";
    }

    function playAlarm() {
        const audio = new Audio("https://cdn.freesound.org/previews/415/415209_5121236-lq.mp3");
        audio.volume = volumeSuono;

        let startTime = Date.now();

        audioInterval = setInterval(() => {
            let now = Date.now();
            if (now - startTime > durataSuono) {
                clearInterval(audioInterval);
                return;
            }
            audio.currentTime = 0;
            audio.play().catch(err => console.warn("Errore nella riproduzione audio:", err));
        }, 1000);
    }

    function cliccaBottone() {
        const elements = document.querySelectorAll("button, input, a, div, span, label");

        for (let el of elements) {
            const testo = normalizeText(el.innerText || el.value || el.textContent);

            if (testo.includes(parolaTarget)) {
                console.log("[Tampermonkey] Trovato elemento con testo che contiene 'prosegui':");
                console.log(el);

                try {
                    el.click();
                    console.log("[Tampermonkey] Bottone cliccato con successo.");
                    playAlarm();
                } catch (e) {
                    console.warn("[Tampermonkey] Errore nel click:", e);
                }

                return true;
            }
        }
        return false;
    }

    const observer = new MutationObserver(() => {
        cliccaBottone();
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });

    window.addEventListener("load", () => {
        setTimeout(cliccaBottone, 1000);
    });

    setInterval(() => {
        console.log("[Tampermonkey] Script attivo...");
    }, 5000);
})();

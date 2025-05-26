// ==UserScript==
// @name         Auto Presenza Teleskill Intelligente (Overlay Persistente)
// @namespace    http://tampermonkey.net/
// @version      1.7
// @description  Clic automatico intelligente su "PROSEGUI" con overlay e suono persistenti fino al click
// @match        *://tlive.teleskill.it/*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const parolaTarget = "prosegui";
    const volumeSuono = 0.7;

    let overlay = null;
    let style = null;
    let audio = null;
    let audioInterval = null;

    function normalizeText(text) {
        return text ? text.toLowerCase().trim() : "";
    }

    function isVisible(el) {
        return el.offsetParent !== null && !el.disabled;
    }

    function containsProsegui(el) {
        const text = normalizeText(el.innerText || el.value || el.textContent);
        return text.includes(parolaTarget);
    }

    function avviaAllarme() {
        if (overlay || audioInterval) return; // già attivo

        // Overlay rosso lampeggiante
        overlay = document.createElement("div");
        overlay.style.position = "fixed";
        overlay.style.top = 0;
        overlay.style.left = 0;
        overlay.style.width = "100vw";
        overlay.style.height = "100vh";
        overlay.style.backgroundColor = "rgba(255, 0, 0, 0.4)";
        overlay.style.zIndex = 999999;
        overlay.style.pointerEvents = "none";
        overlay.style.animation = "blink 1s step-start infinite";
        document.body.appendChild(overlay);

        style = document.createElement("style");
        style.innerHTML = "@keyframes blink { 50% { opacity: 0; } }";
        document.head.appendChild(style);

        // Suono
        audio = new Audio("https://cdn.freesound.org/previews/415/415209_5121236-lq.mp3");
        audio.volume = volumeSuono;

        audioInterval = setInterval(() => {
            audio.currentTime = 0;
            audio.play().catch(err => console.warn("Errore audio:", err));
        }, 1000);
    }

    function fermaAllarme() {
        if (overlay) overlay.remove();
        if (style) style.remove();
        if (audioInterval) clearInterval(audioInterval);

        overlay = null;
        style = null;
        audio = null;
        audioInterval = null;
    }

    function cliccaElemento(el) {
        try {
            el.click();
            console.log("[Tampermonkey] Cliccato:", el);
            fermaAllarme(); // fermiamo l'allarme dopo click
            return true;
        } catch (e) {
            console.warn("[Tampermonkey] Errore nel click:", e);
            return false;
        }
    }

    function cercaEBotta() {
        const cliccabili = document.querySelectorAll("button, input[type='button'], a");

        for (let el of cliccabili) {
            if (!isVisible(el)) continue;
            if (containsProsegui(el)) {
                console.log("[Tampermonkey] Cliccabile con testo 'prosegui':", el);
                avviaAllarme();
                return cliccaElemento(el);
            }
        }

        const tutti = document.querySelectorAll("body *");
        for (let el of tutti) {
            if (!isVisible(el)) continue;
            if (containsProsegui(el)) {
                const figliCliccabili = el.querySelectorAll("button, input[type='button'], a");
                for (let figlio of figliCliccabili) {
                    if (isVisible(figlio)) {
                        console.log("[Tampermonkey] Trovato 'prosegui' come genitore di cliccabile:", el);
                        avviaAllarme();
                        return cliccaElemento(figlio);
                    }
                }
            }
        }

        fermaAllarme(); // se non c'è più nulla, fermiamo l'allarme
        return false;
    }

    const observer = new MutationObserver(() => {
        cercaEBotta();
    });

    observer.observe(document.body, { childList: true, subtree: true });

    window.addEventListener("load", () => {
        setTimeout(cercaEBotta, 1000);
    });

    setInterval(() => {
        console.log("[Tampermonkey] Script attivo...");
    }, 5000);
})();

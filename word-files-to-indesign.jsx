#target "InDesign"

(function () {
    var MAIN_MASTER_NAME = "A"; // Change if your main layout master uses a different label
    var SEPARATOR_MASTER_NAME = "B"; // Change if your separator master uses a different label
    var INSERT_SEPARATOR_AFTER_LAST = false; // Set true if you also want a separator after the final DOCX

    if (app.documents.length === 0) {
        alert("Open the document that should receive the DOCX content before running the script.");
        return;
    }

    var doc = app.activeDocument;

    if (!confirmSetup()) {
        alert("Abgebrochen. Bitte prüfen Sie das Dokument-Setup.");
        return;
    }

    var folder = Folder.selectDialog("Select the folder that contains the DOCX files");
    if (!folder) {
        return;
    }

    var docxFiles = collectDocxFiles(folder);
    if (!docxFiles.length) {
        alert("No DOCX files found in the chosen folder.");
        return;
    }

    var mainMaster = findMasterSpread(doc, MAIN_MASTER_NAME);
    if (!mainMaster) {
        alert('Could not find master spread "' + MAIN_MASTER_NAME + '". Adjust MAIN_MASTER_NAME at the top of the script.');
        return;
    }

    var separatorMaster = findMasterSpread(doc, SEPARATOR_MASTER_NAME);
    if (!separatorMaster) {
        alert('Could not find master spread "' + SEPARATOR_MASTER_NAME + '". Adjust SEPARATOR_MASTER_NAME at the top of the script.');
        return;
    }

    function confirmSetup() {
        var dialog = new Window('dialog', 'Setup prüfen');
        dialog.orientation = 'column';
        dialog.alignChildren = 'fill';
        dialog.spacing = 12;

        var panel = dialog.add('panel', undefined, 'Voraussetzungen');
        panel.orientation = 'column';
        panel.alignChildren = 'left';
        panel.margins = 12;
        panel.spacing = 6;

        var lines = [
            '- Masterseite A enthält eine Doppelseite mit verbundenen leeren Textrahmen.',
            '- Masterseite B besteht aus einer einzelnen Seite als Trenner.',
            '- Alle anderen Textrahmen bleiben unberührt.',
            '- Die ausgewählten DOCX-Dateien werden alphabetisch platziert.'
        ];
        for (var i = 0; i < lines.length; i++) {
            panel.add('statictext', undefined, lines[i]);
        }

        dialog.add('statictext', undefined, 'Ist dieses Setup korrekt?');

        var buttonGroup = dialog.add('group');
        buttonGroup.alignment = 'right';
        buttonGroup.spacing = 12;
        buttonGroup.add('button', undefined, 'Fortfahren', { name: 'ok' });
        buttonGroup.add('button', undefined, 'Abbrechen', { name: 'cancel' });

        return dialog.show() === 1;
    }

    var previousWordPrefs = cloneProperties(app.wordRTFImportPreferences.properties);
    var previousSmartTextReflow = doc.textPreferences.smartTextReflow;
    try {
        configureWordImportPreferences();
        doc.textPreferences.smartTextReflow = true;

        app.doScript(function () {
            processFiles(doc, docxFiles, mainMaster, separatorMaster, INSERT_SEPARATOR_AFTER_LAST);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Place DOCX Files");
    } catch (error) {
        alert(error.message);
    } finally {
        restoreProperties(app.wordRTFImportPreferences, previousWordPrefs);
        doc.textPreferences.smartTextReflow = previousSmartTextReflow;
    }

    function processFiles(documentRef, files, mainMasterSpread, separatorMasterSpread, insertSeparatorAfterLast) {
        for (var i = 0; i < files.length; i++) {
            var file = files[i];
            var newPages = addMasterSpreadPages(documentRef, mainMasterSpread);
            if (!newPages.length) {
                throw new Error("Failed to add pages based on master " + mainMasterSpread.name + " for file " + file.name);
            }

            var targetFrame = null;
            for (var p = 0; p < newPages.length && !targetFrame; p++) {
                targetFrame = findThreadStartFrame(newPages[p]);
            }
            if (!targetFrame) {
                throw new Error("No threaded text frame found on the newly created spread for " + file.name + ".");
            }

            clearFrameText(targetFrame);
            placeDocxIntoFrame(targetFrame, file);

            if (i < files.length - 1 || insertSeparatorAfterLast) {
                addSeparatorPage(documentRef, separatorMasterSpread);
            }
        }
    }

    function collectDocxFiles(sourceFolder) {
        var files = sourceFolder.getFiles(function (entry) {
            return entry instanceof File && /\.docx$/i.test(entry.name);
        });

        files.sort(function (a, b) {
            var nameA = a.name.toLowerCase();
            var nameB = b.name.toLowerCase();
            if (nameA < nameB) {
                return -1;
            }
            if (nameA > nameB) {
                return 1;
            }
            return 0;
        });

        return files;
    }

    function findMasterSpread(documentRef, baseName) {
        var spreads = documentRef.masterSpreads;
        var normalizedBase = baseName.toLowerCase();
        for (var i = 0; i < spreads.length; i++) {
            var candidate = spreads[i];
            var candidateName = candidate.name.toLowerCase();
            if (candidateName === normalizedBase || candidateName.indexOf(normalizedBase + "-") === 0) {
                return candidate;
            }
        }
        return null;
    }

    function addMasterSpreadPages(documentRef, masterSpread) {
        var createdPages = [];
        var masterPages = masterSpread.pages;
        for (var i = 0; i < masterPages.length; i++) {
            var newPage = documentRef.pages.add(LocationOptions.AT_END);
            applyMasterToPage(newPage, masterSpread);
            createdPages.push(newPage);
        }
        return createdPages;
    }

    function applyMasterToPage(page, masterSpread) {
        if (!masterSpread || !masterSpread.isValid || !page || !page.isValid) {
            return;
        }

        try {
            page.appliedMaster = masterSpread;
            return;
        } catch (err) {
            // Fall back to page-specific application below.
        }

        var masterPages = masterSpread.pages;
        if (!masterPages || !masterPages.length) {
            return;
        }

        var target = null;
        var lastIndex = masterPages.length - 1;
        if (page.side === PageSideOptions.LEFT_HAND && masterPages.length > 1) {
            target = masterPages[0];
        } else if (page.side === PageSideOptions.RIGHT_HAND && masterPages.length > 1) {
            target = masterPages[lastIndex];
        } else {
            target = masterPages[0];
        }

        if (target && target.isValid) {
            try {
                page.appliedMaster = target;
            } catch (err2) {
                // Give up silently if the version only accepts full spreads.
            }
        }
    }

    function addSeparatorPage(documentRef, masterSpread) {
        var masterPages = masterSpread.pages;
        if (!masterPages || !masterPages.length) {
            return;
        }

        var targetMasterPage = masterPages[0];
        if (!targetMasterPage || !targetMasterPage.isValid) {
            return;
        }

        var separatorPage = documentRef.pages.add(LocationOptions.AT_END);
        try {
            separatorPage.appliedMaster = targetMasterPage;
        } catch (err) {
            try {
                separatorPage.appliedMaster = masterSpread;
            } catch (err2) {
                // Unable to apply the separator master; leave the page unattached.
            }
        }
    }

    function findThreadStartFrame(page) {
        var frames = page.textFrames;
        for (var i = 0; i < frames.length; i++) {
            var frame = ensurePageFrame(frames[i], page);
            if (!frame) {
                continue;
            }
            var previous = null;
            try {
                previous = frame.previousTextFrame;
            } catch (errPrev) {
                previous = null;
            }
            var next = null;
            try {
                next = frame.nextTextFrame;
            } catch (errNext) {
                next = null;
            }
            var isThreadStart = (!previous || !previous.isValid) && next && next.isValid;
            if (isThreadStart) {
                return frame;
            }
        }
        return frames.length ? ensurePageFrame(frames[0], page) : null;
    }

    function ensurePageFrame(frame, page) {
        if (!frame || !frame.isValid) {
            return null;
        }

        try {
            if (frame.parentPage === page) {
                return frame;
            }
        } catch (err) {
            // Some legacy objects might not expose parentPage (fall through to override).
        }

        if (typeof frame.override === 'function') {
            try {
                return frame.override(page);
            } catch (err2) {
                // Ignore override failures; return null so caller can try the next frame.
                return null;
            }
        }
        return null;
    }

    function getParentStory(item) {
        if (!item || !item.isValid) {
            return null;
        }
        try {
            if (item.parentStory && item.parentStory.isValid) {
                return item.parentStory;
            }
        } catch (err) {
            // Fall through to alternate lookup.
        }
        try {
            if (item.texts && item.texts.length) {
                var firstText = item.texts[0];
                if (firstText && firstText.isValid && firstText.parentStory && firstText.parentStory.isValid) {
                    return firstText.parentStory;
                }
            }
        } catch (err2) {
            // Ignore if not available.
        }
        return null;
    }

    function clearFrameText(frame) {
        if (!frame || !frame.isValid) {
            return;
        }
        var story = getParentStory(frame);
        if (story && story.isValid) {
            story.contents = "";
            return;
        }
        try {
            frame.contents = "";
        } catch (err) {
            // Ignore if the frame cannot be cleared directly.
        }
    }

    function placeDocxIntoFrame(targetFrame, file) {
        var placedItems = targetFrame.place(file);
        var story = null;
        if (placedItems && placedItems.length) {
            story = getParentStory(placedItems[0]);
        }
        if (!story) {
            story = getParentStory(targetFrame);
        }
        if (story && story.isValid) {
            try {
                story.recompose();
            } catch (err) {
                // Some versions recompose automatically; ignore failures.
            }
        }
    }

    function configureWordImportPreferences() {
        var prefs = app.wordRTFImportPreferences;
        safeSet(prefs, 'importStyles', false);
        safeSet(prefs, 'preserveGraphics', true);
        safeSet(prefs, 'convertBulletsAndNumbersToText', false);
        safeSet(prefs, 'preserveTrailingSpaces', true);
    }

    function cloneProperties(sourceProps) {
        var clone = {};
        for (var key in sourceProps) {
            if (sourceProps.hasOwnProperty(key)) {
                clone[key] = sourceProps[key];
            }
        }
        return clone;
    }

    function restoreProperties(target, storedProps) {
        for (var key in storedProps) {
            if (storedProps.hasOwnProperty(key)) {
                try {
                    target[key] = storedProps[key];
                } catch (err) {
                    // Ignore properties that cannot be restored (version differences, etc.).
                }
            }
        }
    }

    function safeSet(target, propertyName, value) {
        if (!target) {
            return;
        }
        try {
            target[propertyName] = value;
        } catch (err) {
            // Property not available or read-only in this version; ignore.
        }
    }

})();

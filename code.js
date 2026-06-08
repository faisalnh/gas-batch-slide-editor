/**
 * Serves the HTML file for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Millennia Slide Deck Editor")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Utility to extract a Google Drive file/folder ID from a URL.
 */
function extractIdFromUrl(url) {
  if (!url) return "";
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : url;
}

// --- CACHE, LOGGING & SUMMARY ---
const CACHE_EXPIRATION_SECONDS = 600; // 10 minutes

function getDefaultJobState() {
  return {
    log: "",
    status: "running",
    summary: {
      processedFiles: 0,
      editedFiles: 0,
      skippedFiles: 0,
      errorFiles: 0,
      totalChanges: 0,
      currentFile: "",
      message: "Preparing job...",
      errors: [],
    },
  };
}

function getJobState(jobId) {
  const cache = CacheService.getScriptCache();
  const currentLog = cache.get(jobId);
  if (!currentLog) return getDefaultJobState();

  const logObject = JSON.parse(currentLog);
  const defaultState = getDefaultJobState();
  logObject.summary = Object.assign(
    {},
    defaultState.summary,
    logObject.summary || {},
  );
  return logObject;
}

function saveJobState(jobId, jobState) {
  CacheService.getScriptCache().put(
    jobId,
    JSON.stringify(jobState),
    CACHE_EXPIRATION_SECONDS,
  );
}

function updateLog(jobId, message, status) {
  const logObject = getJobState(jobId);

  // Keep logs internally for debugging, but the UI displays the summary only.
  const timestamp = new Date().toLocaleTimeString();
  logObject.log = `[${timestamp}] ${message}\n` + logObject.log;

  if (status) {
    logObject.status = status;
  }

  saveJobState(jobId, logObject);
}

function updateJobSummary(jobId, summaryUpdates, status) {
  const logObject = getJobState(jobId);
  logObject.summary = Object.assign({}, logObject.summary, summaryUpdates);

  if (status) {
    logObject.status = status;
  }

  saveJobState(jobId, logObject);
}

function finishJob(jobId, summary) {
  const finalStatus = summary.errorFiles > 0 ? "error" : "completed";
  const message =
    summary.errorFiles > 0
      ? "Completed with errors. Review the error summary below."
      : "Process finished successfully.";

  updateJobSummary(jobId, Object.assign({}, summary, { message }), finalStatus);
  updateLog(jobId, message, finalStatus);
}

/**
 * Client-side poller function. Retrieves the latest summary for a given job ID.
 */
function getLogUpdates(jobId) {
  const cache = CacheService.getScriptCache();
  const logData = cache.get(jobId);
  return logData
    ? JSON.parse(logData)
    : {
        status: "error",
        log: "Job ID not found or expired.",
        summary: Object.assign({}, getDefaultJobState().summary, {
          message: "Job ID not found or expired.",
          errorFiles: 1,
          errors: ["Job ID not found or expired."],
        }),
      };
}

// --- TASK STARTER FUNCTIONS ---
// These are called by the UI to initiate a job. They return a jobId for polling.

function startDeleteTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Delete Slides");
  updateJobSummary(jobId, { message: "Deleting slides..." });
  deleteSlidesInFolder(jobId, formData);

  return { jobId: jobId };
}

function startInsertTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Insert Slides");
  updateJobSummary(jobId, { message: "Inserting slides..." });
  insertSlideFromSourceIntoAllPresentations(jobId, formData);
  return { jobId: jobId };
}

function startReplaceTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Replace Slides");
  updateJobSummary(jobId, { message: "Replacing slides..." });
  replaceSlideInBatch(jobId, formData);
  return { jobId: jobId };
}

function startTextReplaceTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Text Replace");
  updateJobSummary(jobId, { message: "Replacing text..." });
  replaceTextInFolder(jobId, formData);
  return { jobId: jobId };
}

function startFileNameReplaceTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: File Name Replace");
  updateJobSummary(jobId, { message: "Replacing file names..." });
  replaceFileNamesInFolder(jobId, formData);
  return { jobId: jobId };
}

// --- CORE LOGIC FUNCTIONS ---

function deleteSlidesInFolder(jobId, formData) {
  const summary = getDefaultJobState().summary;
  summary.message = "Deleting slides...";

  try {
    const folderId = extractIdFromUrl(formData.deleteFolderUrl);
    const startSlide = parseInt(formData.startSlide, 10);
    const endSlide = parseInt(formData.endSlide, 10);

    if (!folderId || isNaN(startSlide) || isNaN(endSlide)) {
      throw new Error("Invalid input.");
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
    updateLog(jobId, `Starting deletion in folder: ${folder.getName()}`);

    while (files.hasNext()) {
      const file = files.next();
      summary.processedFiles++;
      summary.currentFile = file.getName();

      try {
        const presentation = SlidesApp.openById(file.getId());
        const slides = presentation.getSlides();
        const totalSlides = slides.length;
        let deletedCount = 0;

        if (totalSlides < startSlide) {
          summary.skippedFiles++;
          updateLog(jobId, `Skipped ${file.getName()}: not enough slides.`);
          updateJobSummary(jobId, summary);
          continue;
        }

        for (let i = endSlide - 1; i >= startSlide - 1; i--) {
          if (i < totalSlides) {
            slides[i].remove();
            deletedCount++;
          }
        }

        if (deletedCount > 0) {
          summary.editedFiles++;
          summary.totalChanges += deletedCount;
        } else {
          summary.skippedFiles++;
        }

        updateJobSummary(jobId, summary);
        Utilities.sleep(100);
      } catch (fileError) {
        summary.errorFiles++;
        summary.errors.push(`${file.getName()}: ${fileError.message}`);
        updateLog(jobId, `ERROR in ${file.getName()}: ${fileError.message}`);
        updateJobSummary(jobId, summary);
      }
    }

    finishJob(jobId, summary);
  } catch (e) {
    summary.errorFiles++;
    summary.errors.push(e.message);
    updateJobSummary(
      jobId,
      Object.assign({}, summary, { message: `ERROR: ${e.message}` }),
      "error",
    );
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function insertSlideFromSourceIntoAllPresentations(jobId, formData) {
  const summary = getDefaultJobState().summary;
  summary.message = "Inserting slides...";

  try {
    const sourcePresentationId = extractIdFromUrl(formData.insertSourceUrl);
    const destinationFolderId = extractIdFromUrl(formData.insertDestFolderUrl);
    const sourceSlideStartIndex = parseInt(formData.insertSourceStart, 10);
    const sourceSlideEndIndex = parseInt(formData.insertSourceEnd, 10);
    const destinationSlideIndex =
      parseInt(formData.insertDestPosition, 10) || 0;

    if (
      !sourcePresentationId ||
      !destinationFolderId ||
      isNaN(sourceSlideStartIndex) ||
      isNaN(sourceSlideEndIndex)
    ) {
      throw new Error("Invalid input.");
    }

    const sourcePresentation = SlidesApp.openById(sourcePresentationId);
    const sourceSlides = sourcePresentation.getSlides();
    const folder = DriveApp.getFolderById(destinationFolderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
    const slideCountToInsert = sourceSlideEndIndex - sourceSlideStartIndex + 1;

    updateLog(
      jobId,
      `Source: ${sourcePresentation.getName()}, Destination Folder: ${folder.getName()}`,
    );

    while (files.hasNext()) {
      const file = files.next();
      summary.processedFiles++;
      summary.currentFile = file.getName();

      try {
        const destPpt = SlidesApp.openById(file.getId());

        if (destinationSlideIndex > 0) {
          let insertionIndex = destinationSlideIndex - 1;
          for (
            let i = sourceSlideStartIndex - 1;
            i <= sourceSlideEndIndex - 1;
            i++
          ) {
            destPpt.insertSlide(insertionIndex, sourceSlides[i]);
            insertionIndex++;
          }
        } else {
          for (
            let i = sourceSlideStartIndex - 1;
            i <= sourceSlideEndIndex - 1;
            i++
          ) {
            destPpt.appendSlide(sourceSlides[i]);
          }
        }

        summary.editedFiles++;
        summary.totalChanges += slideCountToInsert;
        updateJobSummary(jobId, summary);
        Utilities.sleep(100);
      } catch (fileError) {
        summary.errorFiles++;
        summary.errors.push(`${file.getName()}: ${fileError.message}`);
        updateLog(jobId, `ERROR in ${file.getName()}: ${fileError.message}`);
        updateJobSummary(jobId, summary);
      }
    }

    finishJob(jobId, summary);
  } catch (e) {
    summary.errorFiles++;
    summary.errors.push(e.message);
    updateJobSummary(
      jobId,
      Object.assign({}, summary, { message: `ERROR: ${e.message}` }),
      "error",
    );
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function replaceTextInFolder(jobId, formData) {
  const summary = getDefaultJobState().summary;
  summary.message = "Replacing text...";

  try {
    const folderId = extractIdFromUrl(formData.textReplaceFolderUrl);
    const findText = formData.findText;
    const replaceWithText = formData.replaceWithText || "";
    const scope = formData.textReplaceScope || "all";
    const slideNumber = parseInt(formData.textReplaceSlideNumber, 10);
    const matchCase =
      formData.matchCase === true || formData.matchCase === "true";

    if (!folderId || !findText) {
      throw new Error("Folder URL and find text are required.");
    }

    if (scope === "specific" && isNaN(slideNumber)) {
      throw new Error(
        "Slide number is required when using specific slide scope.",
      );
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);

    updateLog(
      jobId,
      `Starting text replacement in folder: ${folder.getName()}`,
    );

    while (files.hasNext()) {
      const file = files.next();
      summary.processedFiles++;
      summary.currentFile = file.getName();

      try {
        const presentation = SlidesApp.openById(file.getId());
        const slides = presentation.getSlides();
        let replacementCount = 0;

        if (scope === "specific") {
          if (slideNumber < 1 || slideNumber > slides.length) {
            summary.skippedFiles++;
            updateJobSummary(jobId, summary);
            continue;
          }

          replacementCount += slides[slideNumber - 1].replaceAllText(
            findText,
            replaceWithText,
            matchCase,
          );
        } else {
          for (let i = 0; i < slides.length; i++) {
            replacementCount += slides[i].replaceAllText(
              findText,
              replaceWithText,
              matchCase,
            );
          }
        }

        if (replacementCount > 0) {
          summary.editedFiles++;
          summary.totalChanges += replacementCount;
        } else {
          summary.skippedFiles++;
        }

        updateJobSummary(jobId, summary);
        Utilities.sleep(100);
      } catch (fileError) {
        summary.errorFiles++;
        summary.errors.push(`${file.getName()}: ${fileError.message}`);
        updateLog(jobId, `ERROR in ${file.getName()}: ${fileError.message}`);
        updateJobSummary(jobId, summary);
      }
    }

    finishJob(jobId, summary);
  } catch (e) {
    summary.errorFiles++;
    summary.errors.push(e.message);
    updateJobSummary(
      jobId,
      Object.assign({}, summary, { message: `ERROR: ${e.message}` }),
      "error",
    );
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function replaceFileNamesInFolder(jobId, formData) {
  const summary = getDefaultJobState().summary;
  summary.message = "Replacing file names...";

  try {
    const folderId = extractIdFromUrl(formData.fileNameReplaceFolderUrl);
    const findText = formData.fileNameFindText;
    const replaceWithText = formData.fileNameReplaceWithText || "";
    const matchCase =
      formData.fileNameMatchCase === true ||
      formData.fileNameMatchCase === "true";

    if (!folderId || !findText) {
      throw new Error("Folder URL and file name find text are required.");
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
    const escapedFindText = findText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const searchPattern = new RegExp(escapedFindText, matchCase ? "g" : "gi");

    updateLog(
      jobId,
      `Starting file name replacement in folder: ${folder.getName()}`,
    );

    while (files.hasNext()) {
      const file = files.next();
      const originalName = file.getName();
      summary.processedFiles++;
      summary.currentFile = originalName;

      try {
        const newName = originalName.replace(searchPattern, replaceWithText);

        if (newName !== originalName) {
          file.setName(newName);
          summary.editedFiles++;
          summary.totalChanges++;
          updateLog(jobId, `Renamed ${originalName} to ${newName}`);
        } else {
          summary.skippedFiles++;
        }

        updateJobSummary(jobId, summary);
        Utilities.sleep(100);
      } catch (fileError) {
        summary.errorFiles++;
        summary.errors.push(`${originalName}: ${fileError.message}`);
        updateLog(jobId, `ERROR in ${originalName}: ${fileError.message}`);
        updateJobSummary(jobId, summary);
      }
    }

    finishJob(jobId, summary);
  } catch (e) {
    summary.errorFiles++;
    summary.errors.push(e.message);
    updateJobSummary(
      jobId,
      Object.assign({}, summary, { message: `ERROR: ${e.message}` }),
      "error",
    );
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function replaceSlideInBatch(jobId, formData) {
  const summary = getDefaultJobState().summary;
  summary.message = "Replacing slides...";

  try {
    const replacedFolderId = extractIdFromUrl(formData.replaceTargetFolderUrl);
    const replacingFolderId = extractIdFromUrl(formData.replaceSourceFolderUrl);
    const replacedSlideNumber = parseInt(formData.replaceTargetSlideNum, 10);
    const replacingSlideNumber = parseInt(formData.replaceSourceSlideNum, 10);

    if (
      !replacedFolderId ||
      !replacingFolderId ||
      isNaN(replacedSlideNumber) ||
      isNaN(replacingSlideNumber)
    ) {
      throw new Error("Invalid input.");
    }

    const replacedFolder = DriveApp.getFolderById(replacedFolderId);
    const replacingFolder = DriveApp.getFolderById(replacingFolderId);
    const replacingFilesMap = new Map();
    const replacingFilesIterator = replacingFolder.getFilesByType(
      MimeType.GOOGLE_SLIDES,
    );

    while (replacingFilesIterator.hasNext()) {
      const file = replacingFilesIterator.next();
      replacingFilesMap.set(file.getName(), file.getId());
    }

    const replacedFilesIterator = replacedFolder.getFilesByType(
      MimeType.GOOGLE_SLIDES,
    );
    while (replacedFilesIterator.hasNext()) {
      const replacedFile = replacedFilesIterator.next();
      const fileName = replacedFile.getName();
      summary.processedFiles++;
      summary.currentFile = fileName;

      try {
        if (replacingFilesMap.has(fileName)) {
          const destinationPresentation = SlidesApp.openById(
            replacedFile.getId(),
          );
          const sourcePresentation = SlidesApp.openById(
            replacingFilesMap.get(fileName),
          );
          const oldSlide =
            destinationPresentation.getSlides()[replacedSlideNumber - 1];
          const newSlide =
            sourcePresentation.getSlides()[replacingSlideNumber - 1];

          if (!oldSlide || !newSlide) {
            throw new Error("Requested slide number does not exist.");
          }

          destinationPresentation.insertSlide(
            replacedSlideNumber - 1,
            newSlide,
          );
          oldSlide.remove();
          summary.editedFiles++;
          summary.totalChanges++;
        } else {
          summary.skippedFiles++;
        }

        updateJobSummary(jobId, summary);
        Utilities.sleep(100);
      } catch (fileError) {
        summary.errorFiles++;
        summary.errors.push(`${fileName}: ${fileError.message}`);
        updateLog(jobId, `ERROR in ${fileName}: ${fileError.message}`);
        updateJobSummary(jobId, summary);
      }
    }

    finishJob(jobId, summary);
  } catch (e) {
    summary.errorFiles++;
    summary.errors.push(e.message);
    updateJobSummary(
      jobId,
      Object.assign({}, summary, { message: `ERROR: ${e.message}` }),
      "error",
    );
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

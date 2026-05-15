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

// --- CACHE & LOGGING ---
const CACHE_EXPIRATION_SECONDS = 600; // 10 minutes

function updateLog(jobId, message, status) {
  const cache = CacheService.getScriptCache();
  const currentLog =
    cache.get(jobId) || JSON.stringify({ log: "", status: "running" });
  const logObject = JSON.parse(currentLog);

  // Prepend new messages so the latest is always at the top
  const timestamp = new Date().toLocaleTimeString();
  logObject.log = `[${timestamp}] ${message}\n` + logObject.log;

  if (status) {
    logObject.status = status;
  }

  cache.put(jobId, JSON.stringify(logObject), CACHE_EXPIRATION_SECONDS);
}

/**
 * Client-side poller function. Retrieves the latest log for a given job ID.
 */
function getLogUpdates(jobId) {
  const cache = CacheService.getScriptCache();
  const logData = cache.get(jobId);
  return logData
    ? JSON.parse(logData)
    : { status: "error", log: "Job ID not found or expired." };
}

// --- TASK STARTER FUNCTIONS ---
// These are called by the UI to initiate a job. They return a jobId for polling.

function startDeleteTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Delete Slides");

  // We call the main function. The client will poll for updates using the jobId.
  deleteSlidesInFolder(jobId, formData);

  return { jobId: jobId };
}

function startInsertTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Insert Slides");
  insertSlideFromSourceIntoAllPresentations(jobId, formData);
  return { jobId: jobId };
}

function startReplaceTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Replace Slides");
  replaceSlideInBatch(jobId, formData);
  return { jobId: jobId };
}

function startTextReplaceTask(formData) {
  const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  updateLog(jobId, "Job initiated: Text Replace");
  replaceTextInFolder(jobId, formData);
  return { jobId: jobId };
}

// --- CORE LOGIC FUNCTIONS (Refactored for Real-Time Logging) ---

function deleteSlidesInFolder(jobId, formData) {
  try {
    const folderId = extractIdFromUrl(formData.deleteFolderUrl);
    const startSlide = parseInt(formData.startSlide, 10);
    const endSlide = parseInt(formData.endSlide, 10);

    if (!folderId || isNaN(startSlide) || isNaN(endSlide))
      throw new Error("Invalid input.");

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
    updateLog(jobId, `Starting deletion in folder: ${folder.getName()}`);

    while (files.hasNext()) {
      const file = files.next();
      const presentation = SlidesApp.openById(file.getId());
      const slides = presentation.getSlides();
      const totalSlides = slides.length;
      updateLog(
        jobId,
        `Processing file: ${file.getName()} (${totalSlides} slides).`,
      );

      if (totalSlides < startSlide) {
        updateLog(jobId, `  - Skipping: Not enough slides.`);
        continue;
      }

      for (let i = endSlide - 1; i >= startSlide - 1; i--) {
        if (i < totalSlides) {
          slides[i].remove();
          updateLog(jobId, `  - Deleted slide ${i + 1}.`);
          Utilities.sleep(100); // Small delay to allow logs to feel more real-time and prevent overwhelming APIs
        }
      }
    }
    updateLog(jobId, "Process finished successfully.", "completed");
  } catch (e) {
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function insertSlideFromSourceIntoAllPresentations(jobId, formData) {
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
    )
      throw new Error("Invalid input.");

    const sourcePresentation = SlidesApp.openById(sourcePresentationId);
    const sourceSlides = sourcePresentation.getSlides();
    const folder = DriveApp.getFolderById(destinationFolderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);

    updateLog(
      jobId,
      `Source: ${sourcePresentation.getName()}, Destination Folder: ${folder.getName()}`,
    );

    while (files.hasNext()) {
      const file = files.next();
      const destPpt = SlidesApp.openById(file.getId());
      updateLog(jobId, `Processing file: ${file.getName()}`);

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
      updateLog(jobId, `  - Inserted slides into ${file.getName()}`);
      Utilities.sleep(100);
    }
    updateLog(jobId, "Process finished successfully.", "completed");
  } catch (e) {
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function replaceTextInFolder(jobId, formData) {
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
    updateLog(jobId, `Replacing "${findText}" with "${replaceWithText}"`);

    while (files.hasNext()) {
      const file = files.next();
      const presentation = SlidesApp.openById(file.getId());
      const slides = presentation.getSlides();
      let replacementCount = 0;

      updateLog(
        jobId,
        `Processing file: ${file.getName()} (${slides.length} slides).`,
      );

      if (scope === "specific") {
        if (slideNumber < 1 || slideNumber > slides.length) {
          updateLog(jobId, `  - SKIPPED: Slide ${slideNumber} does not exist.`);
          continue;
        }

        replacementCount += slides[slideNumber - 1].replaceAllText(
          findText,
          replaceWithText,
          matchCase,
        );
        updateLog(
          jobId,
          `  - Replaced ${replacementCount} occurrence(s) on slide ${slideNumber}.`,
        );
      } else {
        for (let i = 0; i < slides.length; i++) {
          const count = slides[i].replaceAllText(
            findText,
            replaceWithText,
            matchCase,
          );
          replacementCount += count;

          if (count > 0) {
            updateLog(
              jobId,
              `  - Slide ${i + 1}: replaced ${count} occurrence(s).`,
            );
          }
        }

        updateLog(
          jobId,
          `  - Total replacements in file: ${replacementCount}.`,
        );
      }

      Utilities.sleep(100);
    }

    updateLog(jobId, "Text replacement finished successfully.", "completed");
  } catch (e) {
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

function replaceSlideInBatch(jobId, formData) {
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
    )
      throw new Error("Invalid input.");

    const replacedFolder = DriveApp.getFolderById(replacedFolderId);
    const replacingFolder = DriveApp.getFolderById(replacingFolderId);
    const replacingFilesMap = new Map();
    const replacingFilesIterator = replacingFolder.getFilesByType(
      MimeType.GOOGLE_SLIDES,
    );

    updateLog(
      jobId,
      `Building map from source folder: ${replacingFolder.getName()}`,
    );
    while (replacingFilesIterator.hasNext()) {
      const file = replacingFilesIterator.next();
      replacingFilesMap.set(file.getName(), file.getId());
    }
    updateLog(
      jobId,
      `Found ${replacingFilesMap.size} source files. Starting replacements...`,
    );

    const replacedFilesIterator = replacedFolder.getFilesByType(
      MimeType.GOOGLE_SLIDES,
    );
    while (replacedFilesIterator.hasNext()) {
      const replacedFile = replacedFilesIterator.next();
      const fileName = replacedFile.getName();
      updateLog(jobId, `Checking target: ${fileName}`);

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

        destinationPresentation.insertSlide(replacedSlideNumber - 1, newSlide);
        oldSlide.remove();
        updateLog(jobId, `  - SUCCESS: Replaced slide in ${fileName}.`);
      } else {
        updateLog(
          jobId,
          `  - SKIPPED: No matching source file for ${fileName}.`,
        );
      }
      Utilities.sleep(100);
    }
    updateLog(jobId, "Process finished successfully.", "completed");
  } catch (e) {
    updateLog(jobId, `ERROR: ${e.message}`, "error");
  }
}

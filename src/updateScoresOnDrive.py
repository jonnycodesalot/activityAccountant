import googleDriveClient as gd
import activityAccountant as aa
import os
import shutil

if __name__ == "__main__":
    # Create the google service and download all hte files we need
    gdService = gd.createService()
    localInputDir = "/tmp/activityAccountant/input"
    if os.path.isdir(localInputDir):
        shutil.rmtree(localInputDir)
    gd.downloadExcelDirectory(
        gdService,
        gd.getFolderIdByName(gdService, "ActivityAccounting"),
        localInputDir,
        # Don't download old scoring results
        ignoreNames=["scoring"],
    )
    # Delete the old output dir
    localOutputDir = "/tmp/activityAccountant/results"
    if os.path.isdir(localOutputDir):
        shutil.rmtree(localOutputDir)
    # Create the accountant - this is where the magic happens
    accountant = aa.Accountant(localInputDir, localOutputDir)
    # Export the results
    resultFilePath = accountant.exportResults()
    # Upload the results to drive
    gd.uploadSpreadsheet(
        gdService,
        gd.getFolderIdByName(gdService, "scoring"),
        resultFilePath,
    )

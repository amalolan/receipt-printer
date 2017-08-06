// Gets the folder object given the path in the format root/parent/child/grandchild and so on.
function getFolder(path) {
    if (path === "null"){
      Logger.log("No path given -- from getFolder()");
      return null;
    }
    var folder = getDriveFolder(path);    
    Logger.log(folder.getUrl());  
    return folder;  
}

function getDriveFolder(path) {
  
  var name, folder, search, fullpath;
  
  // Remove extra slashes and trim the path
  fullpath = path.replace(/^\/*|\/*$/g, '').replace(/^\s*|\s*$/g, '').split("/");
  Logger.log(fullpath);
  
  // Always start with the main Drive folder
  folder = DriveApp.getRootFolder();
  
  for (var subfolder in fullpath) {
    
    name = fullpath[subfolder];
    search = folder.getFoldersByName(name);
    
    // If folder does not exit, create it in the current level
    folder = search.hasNext() ? search.next() : folder.createFolder(name);
    
  }
  
  return folder;
  
}

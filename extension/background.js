chrome.action.onClicked.addListener((tab) => {
  chrome.tabs.create({
    url: "https://script.google.com/macros/s/AKfycbzydV2LXVIQF6p7rJpVmXXKLgEOPhHBlA4fyYa5kRhPy3NYUwsVQnQ8Rz-xqINXsHoZiA/exec"
  });
});

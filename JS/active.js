document.addEventListener("DOMContentLoaded", function() {
    setActiveLink();
});

function setActiveLink() {
    var path = window.location.pathname;
    var homeLink = document.querySelector("#home-link");
    var splitLink = document.querySelector("#split-link");
    var mergeLink = document.querySelector("#merge-link");
    var ocrLink = document.querySelector("#ocr-link");
    var imageLink = document.querySelector("#image-link");
    var docxLink = document.querySelector("#docx-link");

    if (homeLink && path.includes("index.html")) {
        homeLink.classList.add("active");
    } else if (splitLink && path.includes("split.html")) {
        splitLink.classList.add("active");
    }
    else if (mergeLink && path.includes("merge.html")) {
        mergeLink.classList.add("active");
    }
    else if (ocrLink && path.includes("ocr.html")) {
        ocrLink.classList.add("active");
    }
    else if (imageLink && path.includes("image.html")) {
        imageLink.classList.add("active");
    }
    else if (docxLink && path.includes("docx.html")) {
        docxLink.classList.add("active");
    }
}

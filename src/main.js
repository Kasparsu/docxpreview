import createReport from 'docx-templates';
function createAndDownloadBlobFile(body, filename, extension = 'pdf') {
    const blob = new Blob([body]);
    const fileName = `${filename}.${extension}`;
    if (navigator.msSaveBlob) {
        // IE 10+
        navigator.msSaveBlob(blob, fileName);
    } else {
        const link = document.createElement('a');
        // Browsers that support HTML5 download attribute
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', fileName);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }
}
window.onload = () => {
    var mammoth = require("mammoth");
    fetch('template.docx').then(resp => {

        resp.arrayBuffer().then(buf => {
            createReport({
                template: buf,
                data: {name: 'John', surname: 'Appleseed'},
            }).then(report => {

                mammoth.convertToHtml({arrayBuffer: report}).then(function (resultObject) {
                    document.querySelector('#preview').innerHTML = resultObject.value;
                    console.log(resultObject.value)
                })
                createAndDownloadBlobFile(
                    report,
                    'report',
                    'docx'
                );
            });
        });


    });
}
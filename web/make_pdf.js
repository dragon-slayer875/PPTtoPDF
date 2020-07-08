function get_excel() {
    eel.choose_excel()(setexcel)
}
function setexcel(params) {
    document.getElementById("excel_path").innerHTML=params[0];
    document.getElementById("data-display").innerHTML=params[1];
}

function get_ppt() {
    eel.choose_ppt()(setppt)
}
function setppt(path) {
    document.getElementById("ppt_path").innerHTML=path
}
function get_save() {
    eel.choose_save()(setsave)
}
function setsave(params) {
    document.getElementById("save_path").innerHTML=params
}
function make_pdf(params) {
    var excelpath = document.getElementById("excel_path").innerHTML;
    var pptpath = document.getElementById("ppt_path").innerHTML;
    var savepath = document.getElementById("save_path").innerHTML;
    var color = document.getElementById("hex").innerHTML;
    eel.make_pdfs(excelpath, pptpath, savepath, color)
}
window.onload=function(){
    document.getElementById("hex").innerHTML = document.getElementById("color").value
    document.getElementById("color").addEventListener("input", function() {
    document.getElementById("hex").innerHTML = document.getElementById("color").value;
}, false);
}

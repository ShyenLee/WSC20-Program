async function loadExcel(target, file) {
    try {
        const response = await fetch(file);
        if (!response.ok) throw new Error("File not found");
        
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const html = XLSX.utils.sheet_to_html(worksheet);
        
        target.innerHTML = html;
    } catch (err) {
        target.innerHTML = `<p style="color:red;">Failed to load Excel file: ${file}</p>`;
        console.error(err);
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll(".excel-table").forEach(div => {
        const src = div.getAttribute("data-src");
        if (src) loadExcel(div, src);
    });
});

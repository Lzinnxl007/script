import XLSX from "xlsx";

const workbook = XLSX.readFile("leeds.xlsx");
const worksheet = workbook.Sheets["Página1"];

const jsonData = XLSX.utils
  .sheet_to_json(worksheet, {
    raw: false,
    header: 1, // Usar a primeira linha como cabeçalho
    range: 1, // Iniciar a leitura a partir da segunda linha (após o cabeçalho)
    blankrows: false,
    defval: "",
  })
  .filter((row) => {
    return row[1] && !row[1].toString().startsWith("11") && row[0].toString().includes("Turismo" || "turismo" || "viagem" || "viagens" || "travel" || "Travel" || "Agencia" || "agencia" || "Tur" || "tur" || "Viagem");
  });

console.log(jsonData.length);

const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.json_to_sheet(jsonData);

XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Página1");

const filename = "franca"; // Nome do arquivo final

try {
  XLSX.writeFile(newWorkbook, `./${filename}.xlsx`);
  console.log("Arquivo Excel filtrado salvo com sucesso.");
} catch (error) {
  console.error("Erro ao salvar o arquivo Excel filtrado:", error);
}

// Archivo script.js
function getRandomColor() {
    // Generar un valor aleatorio entre 0 y 255 para cada componente de color
    const r = Math.floor(Math.random() * 256);
    const g = Math.floor(Math.random() * 256);
    const b = Math.floor(Math.random() * 256);

    // Retornar el color en formato RGB
    return `rgb(${r}, ${g}, ${b})`;
}
document.getElementById("colorButton").addEventListener("click", function() {
    document.body.style.backgroundColor = getRandomColor();
});
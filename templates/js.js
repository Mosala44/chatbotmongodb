try {
    const response = await fetch("{% url 'generar_informe' %}", {
        method: "POST",
        body: formData
    });

    if (!response.ok) {
        throw new Error("Error al generar el informe");
    }

    // Descargar el archivo generado
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "informe_combinado.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();

} catch (error) {
    console.error("❌ Error:", error);
    alert("Hubo un problema al generar el informe.");
}
};
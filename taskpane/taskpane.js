Office.onReady(() => {
  // Office is ready
});

function addTlpLabel() {
  const tlpValue = document.getElementById("tlpSelect").value;
  Word.run(async (context) => {
    const sections = context.document.sections;
    context.load(sections, "headers");
    await context.sync();

    const header = sections.items[0].getHeader("Primary");
    header.insertText(tlpValue, "Start");

    await context.sync();
  }).catch((error) => {
    console.error("Chyba:", error);
  });
}

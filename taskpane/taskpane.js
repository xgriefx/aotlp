let addTlpLabel;

Office.onReady(() => {
  addTlpLabel = function(tlpValue) {
    Word.run(async (context) => {
      const sections = context.document.sections;
      context.load(sections, "headers");
      await context.sync();

      const header = sections.items[0].getHeader("Primary");

      // Clear existing content in header
      header.clear();

      // Insert new right-aligned label
      const paragraph = header.insertParagraph(tlpValue, "Start");
      paragraph.alignment = "Right";

      await context.sync();
    }).catch((error) => {
      console.error("Chyba:", error);
    });
  };
});

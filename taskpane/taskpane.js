let addTlpLabel;

Office.onReady(() => {
  addTlpLabel = function(tlpValue) {
    Word.run(async (context) => {
      const sections = context.document.sections;
      context.load(sections, "headers");
      await context.sync();

      const header = sections.items[0].getHeader("Primary");
      const paragraphs = header.paragraphs;
      context.load(paragraphs, "items, items/text");

      await context.sync();
      console.log("search");
      // Hľadaj existujúci TLP štítok
      let found = false;
      for (let p of paragraphs.items) {
        if (p.text.trim().startsWith("TLP:")) {
          console.log("Found");
          p.insertText(tlpValue, "Replace");
          p.alignment = "Right";
          p.font.color = "#E8E8E8";
          found = true;
          break;
        }
      }

      // Ak nenašiel, pridaj nový odstavec
      if (!found) {
        const newP = header.insertParagraph(tlpValue, "Start");
        newP.alignment = "Right";
        newP.font.color = "#595959";
      }

      await context.sync();
    }).catch((error) => {
      console.error("Chyba pri vkladaní štítku:", error);
    });
  };
});

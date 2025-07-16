let addTlpLabel;

Office.onReady(() => {
  addTlpLabel = function (tlpValue) {
    Word.run(async (context) => {
      const sections = context.document.sections;
      context.load(sections, "headers");
      await context.sync();

      const header = sections.items[0].getHeader("Primary");
      const paragraphs = header.paragraphs;
      context.load(paragraphs, "items");

      await context.sync();

      let found = false;
      for (let p of paragraphs.items) {
        const range = p.getRange();
        range.load("text");
        await context.sync();

        const cleanText = range.text.trim();

        if (/^TLP:/i.test(cleanText)) {
          range.insertText(tlpValue, Word.InsertLocation.replace);
          range.font.color = "#E8E8E8";
          range.paragraphs.getFirst().alignment = "Right";
          found = true;
          break;
        }
      }

      if (!found) {
        const newP = header.insertParagraph(tlpValue, Word.InsertLocation.end);
        newP.alignment = "Right";
        newP.font.color = "#E8E8E8";
      }

      await context.sync();
    }).catch((error) => {
      console.error("Chyba pri práci s TLP štítkom:", error);
    });
  };
});

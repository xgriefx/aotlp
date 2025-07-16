let addTlpLabel;

Office.onReady(() => {
  addTlpLabel = function(tlpValue) {
    Word.run(async (context) => {
      const sections = context.document.sections;
      const body = context.document.body;
      context.load(sections, "headers");
      context.load(body, "pageWidth");

      await context.sync();

      const pageWidth = body.pageWidth;
      const shapeWidth = 200;
      const leftPos = pageWidth - shapeWidth - 40;

      const header = sections.items[0].getHeader("Primary");
      const shapes = header.shapes;
      context.load(shapes, "items");
      await context.sync();

      // Odstráň starý štítok
      for (let shape of shapes.items) {
        if (shape.title === "TLPLabel") {
          shape.delete();
        }
      }

      // Pridaj nové textové pole na dynamickú pozíciu
      const newShape = header.shapes.addTextBox(tlpValue, {
        height: 30,
        width: shapeWidth,
        left: leftPos,
        top: 0
      });

      newShape.title = "TLPLabel";
      newShape.textFrame.textRange.font.bold = true;
      newShape.textFrame.textRange.font.size = 12;
      newShape.textFrame.textRange.font.name = "Calibri";
      newShape.textFrame.textRange.font.color = "#E8E8E8";
      newShape.alignment = "Right";
      newShape.textFrame.horizontalAlignment = "Right";
      newShape.zOrder = "SendToBack";

      await context.sync();
    }).catch((error) => {
      console.error("Chyba:", error);
    });
  };
});

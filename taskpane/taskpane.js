let addTlpLabel;

Office.onReady(() => {
  addTlpLabel = function(tlpValue) {
    Word.run(async (context) => {
      const sections = context.document.sections;
      context.load(sections, "headers");
      await context.sync();

      const header = sections.items[0].getHeader("Primary");
      const shapes = header.shapes;
      context.load(shapes, "items");

      await context.sync();

      // Remove existing TLP label shape if found
      for (let shape of shapes.items) {
        if (shape.title === "TLPLabel") {
          shape.delete();
        }
      }

      // Insert new text box for the label
      const newShape = header.shapes.addTextBox(tlpValue, {
        height: 30,
        width: 200,
        left: 400,
        top: 0
      });

      newShape.title = "TLPLabel";
      newShape.textFrame.textRange.font.bold = true;
      newShape.textFrame.textRange.font.size = 12;
      newShape.textFrame.textRange.font.name = "Calibri";
      newShape.textFrame.textRange.font.color = "#E8E8E8";
      newShape.alignment = "Right";
      newShape.textFrame.horizontalAlignment = "Right";

      // Send shape behind text
      newShape.zOrder = "SendToBack";

      await context.sync();
    }).catch((error) => {
      console.error("Chyba:", error);
    });
  };
});

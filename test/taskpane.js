Office.onReady(() => {
  // Office is ready
});

function insertCard() {
  Word.run(async (context) => {
    const body = context.document.body;

    body.insertHtml(
      `
      <div style="border: 2px solid #0078D4; border-radius: 8px; padding: 10px; margin: 10px 0; background: #f0f8ff;">
        <h3 style="margin-top: 0;">Information Card</h3>
        <p>This is a pre-built card with styled content. You can customize it as needed.</p>
      </div>
      `,
      Word.InsertLocation.end
    );

    await context.sync();
  });
}

function insertText() {
  Word.run(async (context) => {
    const body = context.document.body;

    body.insertText(
      "This is a block of pre-written text, ready to be used in reports or documentation.",
      Word.InsertLocation.end
    );

    await context.sync();
  });
}

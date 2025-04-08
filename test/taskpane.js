Office.onReady(() => {
  const ids = ["profileName", "profileTitle", "bioInput", "profileImage"];
  ids.forEach((id) => {
    const element = document.getElementById(id);
    if (element) {
      const eventType = id === "profileImage" ? "change" : "input";
      element.addEventListener(eventType, renderPreview);
    }
  });
});

function cropImageToCircle(base64Image, size = 300) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = function () {
      const canvas = document.createElement("canvas");
      canvas.width = size;
      canvas.height = size;
      const ctx = canvas.getContext("2d");

      ctx.beginPath();
      ctx.arc(size / 2, size / 2, size / 2, 0, Math.PI * 2, false);
      ctx.closePath();
      ctx.clip();

      ctx.drawImage(img, 0, 0, size, size);
      resolve(canvas.toDataURL("image/png"));
    };
    img.src = base64Image;
  });
}

function renderPreview() {
  const name = document.getElementById("profileName").value || "Name";
  const title = document.getElementById("profileTitle").value || "Title";
  const bio = document.getElementById("bioInput").value || "Biography goes here...";
  const imageInput = document.getElementById("profileImage");
  const previewContainer = document.getElementById("previewCard");

  const render = (imageTag) => {
    previewContainer.innerHTML = `
      <div class="profile-card">
        <div>${imageTag}</div>
        <div class="profile-text">
          <h2>${name}</h2>
          <p class="title">${title}</p>
          <p>${bio}</p>
        </div>
      </div>
    `;
  };

  if (imageInput.files.length > 0) {
    const reader = new FileReader();
    reader.onload = async function () {
      const cropped = await cropImageToCircle(reader.result, 300);
      const imageTag = `<img src="${cropped}" class="profile-image" />`;
      render(imageTag);
    };
    reader.readAsDataURL(imageInput.files[0]);
  } else {
    render(`<div class="profile-image"></div>`);
  }
}

async function insertEditableProfile2() {
  const name = document.getElementById("profileName").value;
  const title = document.getElementById("profileTitle").value;
  const bio = document.getElementById("bioInput").value;
  const imageInput = document.getElementById("profileImage");

  const insertTextCard = async () => {
    const html = `
      <div style="background-color: #0072c6; color: white; padding: 32px 24px; display: flex; align-items: center; border-radius: 16px; gap: 24px; margin: 20px 0; font-family: 'Segoe UI', Calibri, sans-serif;">
        <div style="width: 120px; height: 120px; border-radius: 50%; background-color: #f0f0f0; border: 1px solid #ccc;"></div>
        <div style="flex: 1;">
          <h2 style="margin: 0; font-size: 24px; font-weight: 700;">${name}</h2>
          <p style="margin: 5px 0 10px; font-size: 18px; font-weight: 500;">${title}</p>
          <p style="font-size: 16px; line-height: 1.5;">${bio}</p>
        </div>
      </div>
    `;
    await Word.run(async (context) => {
      context.document.body.insertHtml(html, Word.InsertLocation.end);
      await context.sync();
    });
  };

  if (imageInput.files.length > 0) {
    const reader = new FileReader();
    reader.onload = async function () {
      const cropped = await cropImageToCircle(reader.result, 300);
      const base64Image = cropped.split(",")[1];

      await Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
        paragraph.select();
        const range = paragraph.getRange("Start");
        const shape = context.document.shapes.addImage(base64Image);
        shape.width = 120;
        shape.height = 120;
        shape.left = 0;
        shape.top = 0;

        const textRange = range.insertHtml(
          `
          <div style="background-color: #0072c6; color: white; padding: 32px 24px; display: flex; align-items: center; border-radius: 16px; gap: 24px; font-family: 'Segoe UI', Calibri, sans-serif;">
            <div style="width: 120px; height: 120px;"></div>
            <div style="flex: 1;">
              <h2 style="margin: 0; font-size: 24px; font-weight: 700;">${name}</h2>
              <p style="margin: 5px 0 10px; font-size: 18px; font-weight: 500;">${title}</p>
              <p style="font-size: 16px; line-height: 1.5;">${bio}</p>
            </div>
          </div>
          `,
          Word.InsertLocation.replace
        );
        await context.sync();
      });
    };
    reader.readAsDataURL(imageInput.files[0]);
  } else {
    await insertTextCard();
  }
}
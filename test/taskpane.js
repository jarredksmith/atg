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
    reader.onload = function () {
      const img = new Image();
      img.onload = function () {
        const canvas = document.createElement("canvas");
        const size = 300;
        canvas.width = size;
        canvas.height = size;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, size, size);
        ctx.drawImage(img, 0, 0, size, size);
        const compressedDataUrl = canvas.toDataURL("image/jpeg", 0.8);
        const imageTag = `<img src="${compressedDataUrl}" class="profile-image" />`;
        render(imageTag);
      };
      img.src = reader.result;
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

  const buildHtml = (imageHtml) => `
    <div style="background-color: #0072c6; color: white; padding: 32px 24px; display: flex; align-items: center; border-radius: 16px; gap: 24px; margin: 20px 0; font-family: 'Segoe UI', Calibri, sans-serif;">
      <div style="flex: 0 0 120px;">${imageHtml}</div>
      <div style="flex: 1;">
        <h2 style="margin: 0; font-size: 24px; font-weight: 700;">${name}</h2>
        <p style="margin: 5px 0 10px; font-size: 18px; font-weight: 500;">${title}</p>
        <p style="font-size: 16px; line-height: 1.5;">${bio}</p>
      </div>
    </div>
  `;

  if (imageInput.files.length > 0) {
    const reader = new FileReader();
    reader.onload = function () {
      const img = new Image();
      img.onload = async function () {
        const canvas = document.createElement("canvas");
        const size = 300;
        canvas.width = size;
        canvas.height = size;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, size, size);
        ctx.drawImage(img, 0, 0, size, size);
        const compressedDataUrl = canvas.toDataURL("image/jpeg", 0.8);
        const imageHtml = `<img src="${compressedDataUrl}" alt="Profile Photo" style="width: 120px; height: 120px; border-radius: 50%; border: 1px solid #ccc; object-fit: cover; background-color: #f0f0f0;" />`;

        await Word.run(async (context) => {
          context.document.body.insertHtml(buildHtml(imageHtml), Word.InsertLocation.end);
          await context.sync();
        });
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(imageInput.files[0]);
  } else {
    const placeholder = `<div style="width: 120px; height: 120px; border-radius: 50%; background-color: #f0f0f0; border: 1px solid #ccc;"></div>`;
    await Word.run(async (context) => {
      context.document.body.insertHtml(buildHtml(placeholder), Word.InsertLocation.end);
      await context.sync();
    });
  }
}
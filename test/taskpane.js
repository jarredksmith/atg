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

    body.insertHtml(
      'Alliance Emissions Monitoring is a service line within the Fugitive Emissions division of Alliance, the national leader in On-site Testing & Monitoring, Compliance Consulting, and Laboratory Testing & Analysis. Alliance is dedicated to delivering unmatched quality, competitive pricing, and the unparalleled convenience of a full-service, one-stop provider. <br><br>Alliance\'s expertise in Emissions Monitoring and Source Testing is enhanced by its industry-leading best practices, a network of more than 50 strategically located offices across North America, and over 110 mobile laboratories. These resources uniquely position Alliance to provide top-tier, cost-effective emissions monitoring and testing solutions for clients like Phillips 66 and others across oil & gas, refining, natural gas, chemical, and petrochemical industries. <br><br>As a company that continuously expands its footprint across compliance-related environmental services, Alliance prides itself on attracting and retaining a talented workforce of over 1,750 employees. By fostering career growth, offering comprehensive training, and cultivating a purpose-driven culture, Alliance empowers its team to deliver exceptional results for clients while building fulfilling, long-term career paths. For over 25 years, Alliance has made a significant impact on the environmental industry through expert management in regulatory compliance. <br><br>Our commitment to being compliance and client-focused has afforded us the opportunity to pioneer tailored, customizable solutions and engineered services that address the unique needs of our clients\' maintenance, operations, and environmental departments.<br><br> Alliance is staffed by industry experts who bring deep experience and a commitment to excellence in leading and training their teams. These leaders ensure a focus on quality and clear communication, fostering a culture of precision and reliability. <br><br>The company\'s approach is built on a centralized framework designed to maximize results and uphold the highest standards in every aspect of service delivery. This framework is underpinned by a commitment to integrating people, process, and technology seamlessly. Experienced field personnel leverage advanced monitoring technologies, enhanced by Alliance\'s proprietary Quality Management Process and software, to deliver accurate and efficient solutions.<br><br> A dedicated Technical Services Team supports these efforts by managing proposals, protocols, and reports, ensuring quality control, responsiveness, and scalability. Through this comprehensive and client-focused methodology, Alliance consistently delivers superior results.',
      Word.InsertLocation.end
    );

    await context.sync();
  });
}

function insertEditableProfile() {
  const name = document.getElementById("nameInput").value || "Name";
  const title = document.getElementById("titleInput").value || "Title";
  const photoFile = document.getElementById("photoInput").files[0];

  if (!photoFile) {
    insertCardHtml(name, title, null);
    return;
  }

  const reader = new FileReader();
  reader.onload = function (event) {
    const imageBase64 = event.target.result;
    insertCardHtml(name, title, imageBase64);
  };
  reader.readAsDataURL(photoFile);
}

function insertCardHtml(name, title, imageBase64) {
  const imageStyle = `
    width: 150px;
    height: 150px;
    border: 2px solid #003B6D;
    border-radius: 50%;
    margin-left: 24px;
    object-fit: cover;
    background-color: #f1f1f1;
  `;

  const profileHtml = `
    <div style="
      display: flex;
      align-items: center;
      justify-content: space-between;
      background-color: #0078D4;
      color: white;
      border-radius: 16px;
      padding: 24px;
      margin: 20px 0;
    ">
      <div style="flex: 1;">
        <p style="margin: 0; font-size: 20px; font-weight: bold;">${escapeHtml(name)}</p>
        <p style="margin: 8px 0 0; font-size: 16px;">${escapeHtml(title)}</p>
        <p style="margin: 12px 0 0; font-size: 16px;">Biography</p>
      </div>
      ${
        imageBase64
          ? `<img src="${imageBase64}" style="${imageStyle}" />`
          : `<div style="${imageStyle}"></div>`
      }
    </div>
  `;

  Word.run(async (context) => {
    context.document.body.insertHtml(profileHtml, Word.InsertLocation.end);
    await context.sync();
  });
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

window.insertEditableProfile = insertEditableProfile;

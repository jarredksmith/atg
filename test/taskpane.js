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
function insertBio() {
  Word.run(async (context) => {
    const body = context.document.body;

    body.insertHtml(
      `
      <div style="width: 100%; max-width: 320px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <!-- Header section with placeholder -->
        <div style="height: 40px; background-color: #e9e9e9; display: flex; justify-content: center; align-items: center; color: #999; font-size: 14px; width: 100%;">
          400 × 150
        </div>
        
        <!-- Name section with dark background -->
        <div style="background-color: #1a1a1a; color: white; padding: 10px; width: 100%;">
          <div style="font-size: 24px; font-weight: 600; margin: 0;">Alex Morgan</div>
        </div>
        
        <!-- Job title section with dark background -->
        <div style="background-color: #1a1a1a; color: #4a9bff; padding: 10px; width: 100%;">
          <div style="font-size: 16px; margin: 0;">UX Designer</div>
        </div>
        
        <!-- Bio section with dark background -->
        <div style="background-color: #1a1a1a; color: #cccccc; padding: 10px; width: 100%;">
          <div style="font-size: 14px; line-height: 1.5; margin: 0;">Passionate about creating intuitive and beautiful user experiences. Specializing in responsive web design and user-centered design principles.</div>
        </div>
        
        <!-- Follow button with border -->
        <div style="border: 1px solid #4a9bff; margin-top: 10px; width: 100%;">
          <div style="padding: 8px; text-align: center; font-size: 14px; background-color: transparent; color: #4a9bff;">FOLLOW</div>
        </div>
        
        <!-- More info button with dark background -->
        <div style="border: 1px solid #333; margin-top: 10px; width: 100%;">
          <div style="padding: 8px; text-align: center; font-size: 14px; background-color: #333; color: white;">MORE INFO</div>
        </div>
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

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
       <div style="width: 320px; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1); font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <div style="height: 150px; background-color: #e9e9e9; display: flex; justify-content: center; align-items: center; color: #999; font-size: 14px;">
            400 × 150
        </div>
        <div style="padding: 20px; background-color: #1a1a1a; color: white; text-align: left;">
            <h2 style="font-size: 24px; font-weight: 600; margin-bottom: 5px; margin-top: 0;">Alex Morgan</h2>
            <p style="font-size: 16px; color: #4a9bff; margin-bottom: 15px; margin-top: 0;">UX Designer</p>
            <p style="font-size: 14px; line-height: 1.5; color: #cccccc; margin-bottom: 20px; margin-top: 0;">Passionate about creating intuitive and beautiful user experiences. Specializing in responsive web design and user-centered design principles.</p>
            
            <div style="display: flex; gap: 15px; margin-top: 20px;">
                <a href="#" title="Twitter" style="color: #6b7280; text-decoration: none;">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M22 4.01C21 4.5 20.02 4.84 19 5C18.08 4.13 16.8 3.66 15.42 3.66C12.76 3.66 10.6 5.84 10.6 8.48C10.6 8.87 10.64 9.24 10.73 9.59C6.9 9.39 3.48 7.44 1.17 4.44C0.76 5.17 0.54 6.01 0.54 6.9C0.54 8.58 1.39 10.07 2.67 10.91C1.88 10.89 1.13 10.67 0.5 10.31C0.5 10.33 0.5 10.35 0.5 10.37C0.5 12.7 2.18 14.63 4.41 15.08C4.02 15.19 3.6 15.24 3.16 15.24C2.87 15.24 2.59 15.21 2.31 15.16C2.88 17.05 4.66 18.43 6.76 18.47C5.1 19.74 2.99 20.52 0.7 20.52C0.32 20.52 0.95 20.5 0 20.47C2.12 21.8 4.68 22.61 7.41 22.61C15.42 22.61 20 15.33 20 9.03C20 8.81 20 8.6 19.99 8.39C21 7.63 21.88 6.66 22.58 5.55C21.67 5.95 20.7 6.21 19.68 6.33C20.73 5.7 21.53 4.71 21.91 3.55C20.94 4.12 19.85 4.54 18.69 4.75C17.75 3.76 16.4 3.16 14.93 3.16C12.1 3.16 9.8 5.46 9.8 8.3C9.8 8.69 9.84 9.06 9.93 9.42C5.7 9.22 1.9 7.15 -0.83 3.96C-1.29 4.7 -1.55 5.57 -1.55 6.5C-1.55 8.25 -0.6 9.8 0.75 10.69C-0.13 10.67 -0.96 10.42 -1.69 10.03C-1.69 10.05 -1.69 10.07 -1.69 10.09C-1.69 12.57 0.16 14.63 2.58 15.08C3 15.19 3.41 15.24 3.83 15.24C3.5 15.24 3.18 15.21 2.86 15.16C3.5 17.18 5.45 18.65 7.76 18.67C5.95 20.03 3.67 20.86 1.17 20.86C0.78 20.86 0.39 20.84 0 20.81C2.34 22.24 5.11 23.09 8.09 23.09C17.73 23.09 23 15.31 23 8.44C23 8.19 23 7.94 22.99 7.69C24.01 6.83 24.91 5.78 25.63 4.59C24.63 5.02 23.58 5.3 22.48 5.44C23.6 4.7 24.47 3.59 24.88 2.29C23.82 2.93 22.65 3.39 21.38 3.64C20.36 2.53 18.87 1.85 17.23 1.85C14.1 1.85 11.57 4.39 11.57 7.5C11.57 7.95 11.62 8.39 11.73 8.81C7.06 8.58 2.93 6.32 0.08 2.86C-0.39 3.71 -0.62 4.69 -0.62 5.73C-0.62 7.69 0.38 9.43 1.84 10.43C0.99 10.4 0.19 10.18 -0.55 9.8C-0.55 9.82 -0.55 9.85 -0.55 9.87C-0.55 12.6 1.32 14.85 3.78 15.37C4.2 15.49 4.63 15.55 5.08 15.55C4.75 15.55 4.42 15.52 4.11 15.45C4.77 17.66 6.84 19.25 9.28 19.29C7.35 20.78 4.94 21.67 2.33 21.67C1.89 21.67 1.46 21.65 1.03 21.6C3.5 23.18 6.43 24.1 9.59 24.1C20.2 24.1 26 15.57 26 8.2C26 7.95 26 7.7 25.99 7.46C27.02 6.5 27.92 5.35 28.63 4.04L22 4.01Z" fill="currentColor"/>
                    </svg>
                </a>
                <a href="#" title="Email" style="color: #6b7280; text-decoration: none;">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M20 4H4C2.9 4 2.01 4.9 2.01 6L2 18C2 19.1 2.9 20 4 20H20C21.1 20 22 19.1 22 18V6C22 4.9 21.1 4 20 4ZM20 8L12 13L4 8V6L12 11L20 6V8Z" fill="currentColor"/>
                    </svg>
                </a>
                <a href="#" title="LinkedIn" style="color: #6b7280; text-decoration: none;">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M19 3H5C3.9 3 3 3.9 3 5V19C3 20.1 3.9 21 5 21H19C20.1 21 21 20.1 21 19V5C21 3.9 20.1 3 19 3ZM9 17H6.5V10H9V17ZM7.7 8.7C6.9 8.7 6.3 8.1 6.3 7.3C6.3 6.5 6.9 5.9 7.7 5.9C8.5 5.9 9.1 6.5 9.1 7.3C9.1 8.1 8.5 8.7 7.7 8.7ZM18 17H15.5V13.2C15.5 10.3 12.5 10.6 12.5 13.2V17H10V10H12.5V11.3C13.4 9.5 18 9.3 18 13.2V17Z" fill="currentColor"/>
                    </svg>
                </a>
            </div>
            
            <div style="display: flex; gap: 10px; margin-top: 20px;">
                <div style="padding: 8px 12px; border: 1px solid #4a9bff; border-radius: 4px; text-align: center; font-size: 14px; cursor: pointer; background-color: transparent; color: #4a9bff;">FOLLOW</div>
                <div style="padding: 8px 12px; border: 1px solid #333; border-radius: 4px; text-align: center; font-size: 14px; cursor: pointer; background-color: #333; color: #fff;">MORE INFO</div>
            </div>
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

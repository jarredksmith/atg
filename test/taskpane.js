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
function insertBioCard() {
  Word.run(async (context) => {
    // Insert a paragraph to ensure we have space
    const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
    await context.sync();
    
    // Insert a table to control layout (3 rows x 1 column)
    const table = paragraph.insertTable(4, 1, Word.InsertLocation.after);
    await context.sync();
    
    // Format the table
    table.width = 320;
    table.getBorder("All").type = Word.BorderType.single;
    table.getBorder("All").color = "#CCCCCC";
    table.style = "Table Grid";
    table.styleFirstColumn = false;
    table.styleLastColumn = false;
    table.styleBandedRows = false;
    table.styleBandedColumns = false;
    table.styleLastRow = false;
    table.styleFirstRow = false;
    
    // Get cells for each section
    const headerCell = table.getCell(0, 0);
    const nameCell = table.getCell(1, 0);
    const bioCell = table.getCell(2, 0);
    const buttonsCell = table.getCell(3, 0);
    
    // Format the header/image placeholder cell
    headerCell.shading.color = "#E9E9E9";
    headerCell.body.insertParagraph("400 × 150", Word.InsertLocation.start).alignment = Word.Alignment.centered;
    headerCell.height = 40;
    
    // Format the name/title cell
    nameCell.shading.color = "#1A1A1A";
    const nameRange = nameCell.body.insertParagraph("Alex Morgan", Word.InsertLocation.start).font;
    nameRange.color = "white";
    nameRange.bold = true;
    nameRange.size = 16;
    
    const titleRange = nameCell.body.insertParagraph("UX Designer", Word.InsertLocation.end).font;
    titleRange.color = "#4A9BFF";
    titleRange.size = 12;
    
    // Format the bio cell
    bioCell.shading.color = "#1A1A1A";
    const bioRange = bioCell.body.insertParagraph(
      "Passionate about creating intuitive and beautiful user experiences. Specializing in responsive web design and user-centered design principles.",
      Word.InsertLocation.start
    ).font;
    bioRange.color = "#CCCCCC";
    bioRange.size = 11;
    
    // Add buttons in the last cell using a nested table
    const buttonsTable = buttonsCell.body.insertTable(2, 1, Word.InsertLocation.start);
    buttonsTable.width = 316;
    buttonsTable.getBorder("All").type = Word.BorderType.single;
    buttonsTable.getBorder("All").color = "#CCCCCC";
    
    // Format the FOLLOW button
    const followCell = buttonsTable.getCell(0, 0);
    followCell.getBorder("Bottom").type = Word.BorderType.single;
    followCell.getBorder("Bottom").color = "#CCCCCC";
    const followText = followCell.body.insertParagraph("FOLLOW", Word.InsertLocation.start);
    followText.alignment = Word.Alignment.centered;
    followText.font.color = "#4A9BFF";
    followText.font.size = 11;
    
    // Format the MORE INFO button
    const moreInfoCell = buttonsTable.getCell(1, 0);
    moreInfoCell.shading.color = "#333333";
    const moreInfoText = moreInfoCell.body.insertParagraph("MORE INFO", Word.InsertLocation.start);
    moreInfoText.alignment = Word.Alignment.centered;
    moreInfoText.font.color = "white";
    moreInfoText.font.size = 11;
    
    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
  });
}
function insertProfileCardFromOOXML() {
  Word.run(async (context) => {
    // Insert a clean profile card using OOXML
    // This OOXML is simplified from your extracted document
    context.document.body.insertOoxml(`
      <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
          <pkg:xmlData>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <!-- Blue rectangle background -->
                <w:p>
                  <w:r>
                    <w:rPr><w:noProof/></w:rPr>
                    <w:drawing>
                      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="column"><wp:posOffset>-228600</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="paragraph"><wp:posOffset>-234043</wp:posOffset></wp:positionV>
                        <wp:extent cx="5426075" cy="2198370"/>
                        <wp:wrapNone/>
                        <wp:docPr id="1" name="Rectangle"/>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                              <wps:spPr>
                                <a:xfrm><a:off x="0" y="0"/><a:ext cx="5426075" cy="2198370"/></a:xfrm>
                                <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 7506"/></a:avLst></a:prstGeom>
                                <a:solidFill><a:srgbClr val="156082"/></a:solidFill>
                              </wps:spPr>
                            </wps:wsp>
                          </a:graphicData>
                        </a:graphic>
                      </wp:anchor>
                    </w:drawing>
                  </w:r>
                </w:p>
                
                <!-- Text content box -->
                <w:p>
                  <w:r>
                    <w:rPr><w:noProof/></w:rPr>
                    <w:drawing>
                      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251661312" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="column"><wp:posOffset>2722</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="paragraph"><wp:posOffset>363</wp:posOffset></wp:positionV>
                        <wp:extent cx="2360930" cy="1404620"/>
                        <wp:wrapSquare wrapText="bothSides"/>
                        <wp:docPr id="2" name="Text Box"/>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                              <wps:cNvSpPr txBox="1"/>
                              <wps:spPr>
                                <a:xfrm><a:off x="0" y="0"/><a:ext cx="2360930" cy="1404620"/></a:xfrm>
                                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                              </wps:spPr>
                              <wps:txbx>
                                <w:txbxContent>
                                  <w:p>
                                    <w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:pPr>
                                    <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>Alex Morgan</w:t></w:r>
                                  </w:p>
                                  <w:p>
                                    <w:r><w:rPr><w:color w:val="4a9bff"/></w:rPr><w:t>UX Designer</w:t></w:r>
                                  </w:p>
                                  <w:p>
                                    <w:r><w:t>Passionate about creating intuitive and beautiful user experiences. Specializing in responsive web design and user-centered design principles.</w:t></w:r>
                                  </w:p>
                                </w:txbxContent>
                              </wps:txbx>
                            </wps:wsp>
                          </a:graphicData>
                        </a:graphic>
                      </wp:anchor>
                    </w:drawing>
                  </w:r>
                </w:p>

                <!-- FOLLOW button -->
                <w:p>
                  <w:pPr>
                    <w:jc w:val="center"/>
                    <w:pBdr>
                      <w:top w:val="single" w:sz="4" w:space="1" w:color="4a9bff"/>
                      <w:left w:val="single" w:sz="4" w:space="4" w:color="4a9bff"/>
                      <w:bottom w:val="single" w:sz="4" w:space="1" w:color="4a9bff"/>
                      <w:right w:val="single" w:sz="4" w:space="4" w:color="4a9bff"/>
                    </w:pBdr>
                  </w:pPr>
                  <w:r>
                    <w:rPr><w:color w:val="4a9bff"/></w:rPr>
                    <w:t>FOLLOW</w:t>
                  </w:r>
                </w:p>

                <!-- MORE INFO button -->
                <w:p>
                  <w:pPr>
                    <w:jc w:val="center"/>
                    <w:pBdr>
                      <w:top w:val="single" w:sz="4" w:space="1" w:color="333333"/>
                      <w:left w:val="single" w:sz="4" w:space="4" w:color="333333"/>
                      <w:bottom w:val="single" w:sz="4" w:space="1" w:color="333333"/>
                      <w:right w:val="single" w:sz="4" w:space="4" w:color="333333"/>
                    </w:pBdr>
                    <w:shd w:val="clear" w:color="auto" w:fill="333333"/>
                  </w:pPr>
                  <w:r>
                    <w:rPr><w:color w:val="FFFFFF"/></w:rPr>
                    <w:t>MORE INFO</w:t>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
          </pkg:xmlData>
        </pkg:part>
      </pkg:package>
    `, Word.InsertLocation.end);

    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + JSON.stringify(error));
  });
}
// Alternative approach using shapes and textboxes
function insertBioCardWithShapes() {
  Word.run(async (context) => {
    // Insert a paragraph to ensure we have space
    const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
    await context.sync();
    
    // Create the main container
    const contentControl = paragraph.insertContentControl();
    contentControl.tag = "ProfileCard";
    contentControl.title = "Profile Card";
    contentControl.appearance = Word.ContentControlAppearance.boundingBox;
    contentControl.color = "#CCCCCC";
    contentControl.width = 320;
    await context.sync();
    
    // Add the placeholder section
    const placeholderPara = contentControl.insertParagraph("", Word.InsertLocation.start);
    placeholderPara.insertText("400 × 150", Word.InsertLocation.start);
    placeholderPara.alignment = Word.Alignment.centered;
    placeholderPara.shading.color = "#E9E9E9";
    await context.sync();
    
    // Add the name section with dark background
    const namePara = contentControl.insertParagraph("", Word.InsertLocation.end);
    namePara.shading.color = "#1A1A1A";
    const nameText = namePara.insertText("Alex Morgan", Word.InsertLocation.start);
    nameText.font.color = "white";
    nameText.font.bold = true;
    nameText.font.size = 16;
    await context.sync();
    
    // Add the title section
    const titlePara = contentControl.insertParagraph("", Word.InsertLocation.end);
    titlePara.shading.color = "#1A1A1A";
    const titleText = titlePara.insertText("UX Designer", Word.InsertLocation.start);
    titleText.font.color = "#4A9BFF";
    titleText.font.size = 12;
    await context.sync();
    
    // Add the bio section
    const bioPara = contentControl.insertParagraph("", Word.InsertLocation.end);
    bioPara.shading.color = "#1A1A1A";
    const bioText = bioPara.insertText(
      "Passionate about creating intuitive and beautiful user experiences. Specializing in responsive web design and user-centered design principles.",
      Word.InsertLocation.start
    );
    bioText.font.color = "#CCCCCC";
    bioText.font.size = 11;
    await context.sync();
    
    // Add the FOLLOW button
    const followPara = contentControl.insertParagraph("", Word.InsertLocation.end);
    const followCC = followPara.insertContentControl();
    followCC.title = "Follow Button";
    followCC.tag = "FollowButton";
    followCC.appearance = Word.ContentControlAppearance.boundingBox;
    followCC.color = "#4A9BFF";
    followCC.insertText("FOLLOW", Word.InsertLocation.start);
    followPara.alignment = Word.Alignment.centered;
    followPara.font.color = "#4A9BFF";
    followPara.font.size = 11;
    await context.sync();
    
    // Add the MORE INFO button
    const moreInfoPara = contentControl.insertParagraph("", Word.InsertLocation.end);
    const moreInfoCC = moreInfoPara.insertContentControl();
    moreInfoCC.title = "More Info Button";
    moreInfoCC.tag = "MoreInfoButton";
    moreInfoCC.appearance = Word.ContentControlAppearance.boundingBox;
    moreInfoCC.color = "#333333";
    moreInfoCC.insertText("MORE INFO", Word.InsertLocation.start);
    moreInfoPara.alignment = Word.Alignment.centered;
    moreInfoPara.font.color = "white";
    moreInfoPara.font.size = 11;
    moreInfoPara.shading.color = "#333333";
    
    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
  });
}
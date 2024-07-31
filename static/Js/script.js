
document.getElementById('uploadForm').addEventListener('submit', function (event) {
    event.preventDefault();
    let formData;

    document.getElementById('loadingContainer').style.display = 'block';
    document.getElementById('overlay').style.display = 'block';

    formData = new FormData(this);

    fetch('/process', {
        method: 'POST',
        body: formData
    })
        .then(response => { return response.json() })
        .then(data => {

            // Hide the loading animation and overlay
            document.getElementById('loadingContainer').style.display = 'none';
            document.getElementById('overlay').style.display = 'none';


            // Hide the loading animation
            let ele = document.getElementById('select');
            if (ele.value === "deloitte") {
                //invoke render method with 
                document.getElementById('formDataDisplay').innerHTML = deloitteTemplate;
                document.getElementById('full-name').textContent = data.personal_info.full_name;
                document.getElementById('role-designation').textContent = data.personal_info.role_designation;
                document.getElementById('profile').textContent = data.profile;

                const keySkillsList = document.getElementById('key-skills');
                keySkillsList.innerHTML = '';
                data.skills.key_skills.forEach(skill => {
                    const li = document.createElement('li');
                    li.textContent = skill;
                    keySkillsList.appendChild(li);
                });

                const educationTable = document.getElementById('education');
                educationTable.innerHTML = '';

                // Assuming data.education is a string with formatted educational details
                if (data.education) {
                    // Split the string into individual educational entries
                    const educationEntries = data.education.split(';').map(entry => entry.trim());

                    // Create table rows for each educational entry
                    educationEntries.forEach(entry => {
                        if (entry.length > 0) {
                            const tr = document.createElement('tr');
                            const td = document.createElement('td');
                            td.textContent = entry;
                            tr.appendChild(td);
                            educationTable.appendChild(tr);
                        }
                    });
                }

                const projectsDiv = document.getElementById('projects');
                projectsDiv.innerHTML = '';
                let count = 1;
                data.projects.forEach((proj) => {

                    const projectDiv = document.createElement('div');
                    projectDiv.classList.add('project');
                    projectDiv.innerHTML = `
                    <h3 ><b>Project Name #${count}</b></h3>
                    <table>
                        <tr>
                            <th>Client Name</th>
                            <td>${proj.client_name}</td>
                        </tr>
                        <tr>
                            <th>Duration</th>
                            <td>${proj.duration}</td>
                        </tr>
                        <tr>
                            <th>Role</th>
                            <td>${proj.role}</td>
                        </tr>
                        <tr>
                            <th>Environment</th>
                            <td>${proj.environment}</td>
                        </tr>
                        <tr>
                            <th>Description</th>
                            <td>${proj.description}</td>
                        </tr>
                    </table>
                `;
                    projectsDiv.appendChild(projectDiv);
                    count++;
                });

                // Function to split text by sentences and insert <br> tags
                function formatProfileText(profileText) {
                    // Split text by periods followed by spaces (and optionally more spaces or new lines)
                    const sentences = profileText.split(/(?<=\.)\s+/);
                    // Join sentences with <br> tags
                    return sentences.join('<br>');
                }

                // Assuming `data.profile` is the profile text fetched from the backend
                const formattedProfile = formatProfileText(data.profile);

                // Set the formatted profile text as HTML
                document.getElementById('profile').innerHTML = formattedProfile;

            }
            else {
                // Populate the fields with the parsed resume data
                document.getElementById('name').textContent = data.Name || '';
                document.getElementById('email').textContent = data.Email || '';
                document.getElementById('role').textContent = data.Role || '';
                document.getElementById('profileSummary').textContent = data.Profile || '';


                // Function to split text by sentences while preserving full stops in numbers
                function formatProfileText(profileText) {
                    // Split text by periods followed by spaces, but ignore full stops in numbers
                    const sentences = profileText.split(/(?<=\.)\s+(?!\d)/);
                    // Create a <ul> element for the sentences
                    const profileList = document.createElement('ul');

                    sentences.forEach(sentence => {
                        const listItem = document.createElement('li');
                        listItem.textContent = sentence; // Keep the full stop in the sentence
                        profileList.appendChild(listItem);
                    });

                    return profileList; // Return the <ul> element
                }

                // Populate profile summary
                const profileSummaryContainer = document.getElementById('profileSummary');
                profileSummaryContainer.innerHTML = ''; // Clear previous content

                if (data.Profile) {
                    const formattedProfile = formatProfileText(data.Profile);
                    profileSummaryContainer.appendChild(formattedProfile); // Append the <ul> to the container
                }



                // Populate professional experience
                const professionalExperienceTable = document.getElementById('professionalExperienceTable');
                if (data['Professional Experience'] && data['Professional Experience'].length > 0) {
                    professionalExperienceTable.innerHTML = `
                    <tr>
                        <th>Duration</th>
                        <th>Organization</th>
                        <th>Designation</th>
                    </tr>`;
                    data['Professional Experience'].forEach(exp => {
                        professionalExperienceTable.innerHTML += `
                        <tr>
                            <td>${exp.Duration || ''}</td>
                            <td>${exp.Organization || ''}</td>
                            <td>${exp.Designation || ''}</td>
                        </tr>`;
                    });
                } else {
                    professionalExperienceTable.innerHTML = ''; // Clear table if no data
                    professionalExperienceTable.style.display = 'none';
                    document.getElementById('professionalExperienceHeading').style.display = 'none';
                }

                // Populate technical skills
                const technicalSkillsTable = document.getElementById('technicalSkillsTable');
                if (data['Technical Skills'] && data['Technical Skills'].length > 0) {
                    technicalSkillsTable.innerHTML = `
                    <tr>
                        <th>Expertise</th>
                        <th>Skills</th>
                    </tr>`;
                    data['Technical Skills'].forEach(skill => {
                        technicalSkillsTable.innerHTML += `
                        <tr>
                            <td>${skill.Category || ''}</td>
                            <td>${skill.Skills || ''}</td>
                        </tr>`;
                    });
                } else {
                    technicalSkillsTable.innerHTML = ''; // Clear table if no data
                    technicalSkillsTable.style.display = 'none';
                }

                // Helper function to split text into sentences while preserving full stops in numbers
                function splitIntoSentences(text) {
                    if (typeof text !== 'string') {
                        return []; // Return an empty array if text is not a string
                    }
                    // Split text by periods followed by spaces, but ignore full stops in numbers
                    return text.split(/(?<=\.)\s+(?!\d)/).map(sentence => sentence.trim()).filter(sentence => sentence.length > 0);
                }

                // Populate projects
                const projectsContainer = document.getElementById('projects');
                projectsContainer.innerHTML = ''; // Clear previous content

                if (data['Project Experience - Detail'] && data['Project Experience - Detail'].length > 0) {
                    data['Project Experience - Detail'].forEach((project, index) => {
                        const projectTable = document.createElement('table');
                        projectTable.innerHTML = `
                                <tr>
                                    <th>Project Type</th>
                                    <td class="bold-orange">${project['Project Type'] || ''}</td>
                                </tr>
                                <tr>
                                    <th>Description</th>
                                    <td>${project.Description || ''}</td>
                                </tr>
                                <tr>
                                    <th>Role</th>
                                    <td>${project.Role || ''}</td>
                                </tr>
                                <tr>
                                    <th>Team Size</th>
                                    <td>${project['Team Size'] || ''}</td>
                                </tr>
                                <tr>
                                    <th>Duration</th>
                                    <td>${project.Duration || ''}</td>
                                </tr>
                                <tr>
                                    <th>Tools & Technology</th>
                                    <td>${project['Tools & Technology'] || ''}</td>
                                </tr>
                                <tr>
                                    <th>Responsibilities</th>
                                    <td>
                                        <ul>${splitIntoSentences(project.Responsibilities || '').map(sentence => `<li>${sentence}</li>`).join('')}</ul>
                                    </td>
                                </tr>`;

                        // Add a heading for each project
                        const projectHeading = document.createElement('h3');
                        projectHeading.textContent = `Project ${index + 1}`;
                        projectsContainer.appendChild(projectHeading);
                        projectsContainer.appendChild(projectTable);
                    });
                } else {
                    projectsContainer.style.display = 'none';
                }


                // Populate education
                const educationTable = document.getElementById('educationTable');
                if (data['Educational Details'] && data['Educational Details'].length > 0) {
                    educationTable.innerHTML = `
                    <tr>
                        <th>Qualification</th>
                        <th>University/Board</th>
                        <th>Year</th>
                        <th>Percentage</th>
                    </tr>`;
                    data['Educational Details'].forEach(edu => {
                        educationTable.innerHTML += `
                        <tr>
                            <td>${edu.Qualification || ''}</td>
                            <td>${edu['University/Board'] || ''}</td>
                            <td>${edu.Year || ''}</td>
                            <td>${edu.Percentage || ''}</td>
                        </tr>`;
                    });
                } else {
                    educationTable.style.display = 'none';
                    document.getElementById('educationHeading').style.display = 'none';
                }

                // Populate certifications
                document.getElementById('certifications').textContent = data.Certifications || '';
            }

            document.getElementById('loadingContainer').style.display = 'none';

            if (data.error) {
                alert(data.error);
            }



            else {


                // Show the form data display section
                document.getElementById('uploadFormWrapper').style.display = 'none';
                document.getElementById('formDataDisplay').style.display = 'block';

                const downloadSection = document.getElementById('downloadSection');
                downloadSection.innerHTML = `
                    <!-- Download Button Section -->
                    <div id="downloadSection" style="display: flex; align-items: center; gap: 10px;">
                        <label for="filenameInput" style="margin-right: 10px;">Enter Filename:</label>
                        <input type="text" id="filenameInput" placeholder="Enter filename here" style="flex: 1;">
                        <button id="downloadButton">Download</button>
                    </div>

                `;
                const downloadButton = document.getElementById('downloadButton');
                downloadButton.style.display = 'block'; // Ensure it's visible
                downloadButton.addEventListener('click', exportToWord); // Attach event listener

            }
        })
        .catch(error => {

            // Hide the loading animation and overlay
            document.getElementById('loadingContainer').style.display = 'none';
            document.getElementById('overlay').style.display = 'none';


            console.error('Error:', error);
            alert('An error occurred. Please try again.');
        });
});



const deloitteTemplate = `
    <style>
        body {
            font-family: Calibri, Arial, sans-serif;
            line-height: 1.6;
            background-color: #fbc21e;
            font-size: 14.5px;
            margin: 0;
        }

        table {
            border-collapse: collapse;
            width: 100%;
        }

        th, td {
            padding: 1px;
            background-color: #fff; 
            text-align: left;
            border: 3px solid #000; /* Thin black border for table cells */
            vertical-align: middle; /* Vertically align text to the middle */
        }

        th {
            background-color: #fff; /* White background for table headers */
            font-weight: bold;
            border: 1px solid #000; /* Thin black border for table cells */
        }

        h2, h3 {
            margin: 0;
            padding: 0;
            font-size: 14.5px; /* Set font size for headings */
            text-decoration: none;
        }

        /* Add spacing between section headings and tables */
        section h2 {
            margin-bottom: 10px; /* Adjust this value to increase or decrease the distance */
        }

        /* Add spacing specifically for project headings */
        .project h3 {
            margin-bottom: 10px; /* Adjust this value to increase or decrease the distance */
        }

        section {
            margin-bottom: 20px;
        }

        .project {
            margin-bottom: 20px; /* Add space between each project */
            vertical-align: middle; /* Vertically align text to the middle */
        }

        ul {
            margin: 0;
            padding: 0;
            
            margin-left: 20px;
        }

        .header-bar {
            height: 4px; /* Thickness of the bar */
            background-color: #000; /* Black color for the bar */
            margin-bottom: 20px; /* Space between bar and section */
        }
          
    </style>
    <div >
            <!-- Header Bar -->
             <div align="left">
                <img src="data:image/jpeg;base64,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" alt="Company Logo" class="logo" height="50" width="130">
             </div>
         
            <div class="header-bar" style="height:4px; background-color: #000; margin-bottom: 20px;color:#000; font-size:2px"><p>Hello</p></div>
            
    
            <!-- Personal Information Section -->
            <section class="personal-info">
                <table>
                    <tr>
                        <th>Full Name</th>
                        <td id="full-name"></td>
                    </tr>
                    <tr>
                        <th>Role / Designation:</th>
                        <td id="role-designation"></td>
                    </tr>
                </table>
                  <br>
            </section>
           
            <!-- Profile Section -->
            <section class="profile">
                <h2>Profile:</h2>
                  <br>
                <table>
                    <tr>
                        <td id="profile"></td>
                    </tr>
                </table>
            </section>
              <br>
            <!-- Technical Skills Section -->
            <section class="skills">
                <h2>Technical or Key Skills:</h2>
                  <br>
                <table>
                    <tr>
                        <td>
                            <h2>Key Skills:</h2>
                            <ul id="key-skills">
                                <!-- Skills will be dynamically inserted here -->
                            </ul>
                        </td>
                    </tr>
                </table>
            </section>
              <br>
            <!-- Education Section -->
            <section class="education">
                <h2>Education:</h2>
                  <br>
                <table id="education">
                    <!-- Education entries will be dynamically inserted here -->
                </table>
            </section>

           <br>
                <h2>Experience:</h2>
              <!-- Project Section -->
                   <section class="projects">
                        <div id="projects">
                         <!-- Projects will be dynamically inserted here -->
                       </div>
                   </section>
            
        </div>
`;

// Update loading message function
const messages = [
    "Fetching data... This might take a few seconds.",
    "Analyzing the text for insights... Hang tight, we're on it.",
    "Almost there! Extracting details from your resume.",
    "Getting your data ready... Please be patient."
];

let currentIndex = 0;

function updateMessage() {
    currentIndex = (currentIndex + 1) % messages.length;
    document.getElementById('loadingMessage').innerText = messages[currentIndex];
}

// Update the message every 3 seconds
setInterval(updateMessage, 2000);

// Function to show the filename input field and confirm button
function showFilenameInput() {
    const filenameSection = document.getElementById('filenameSection');
    filenameSection.style.display = 'block';
}

// Function to handle filename confirmation and export to Word
function exportToWord() {
    // Get HTML content from formDataDisplay section
    const formDataDisplayHTML = document.getElementById('formDataDisplay').innerHTML;

    // Create a Blob containing the HTML data
    const blob = new Blob([wrapHtmlInWordTemplate(formDataDisplayHTML)], { type: 'application/msword' });

    // Get the custom filename from the input field
    const customFilename = document.getElementById('filenameInput').value.trim();
    const fileName = customFilename ? customFilename : 'PIT';

    // Create a link element to trigger the download
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${fileName}.doc`; // Use the name for the filename
    document.body.appendChild(link);

    // Trigger the download
    link.click();

    // Clean up
    document.body.removeChild(link);

    // Hide the filename input section after download
    document.getElementById('filenameSection').style.display = 'none';
}

// Event listener for the download button to show the filename input section
document.getElementById('downloadButton').addEventListener('click', showFilenameInput);

// Event listener for the confirm download button to trigger the export
document.getElementById('confirmDownloadButton').addEventListener('click', exportToWord);

// Ensure the script runs after the DOM is fully loaded
document.addEventListener('DOMContentLoaded', (event) => {
    document.getElementById('exportButton').addEventListener('click', showFilenameInput);
});


// Helper function to wrap HTML in a Word-compatible template
function wrapHtmlInWordTemplate(html) {
    // Define a basic Word document structure with inline CSS styles
    const template = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Resume</title>
    <style>
        body {
            font-family: Calibri, Arial, sans-serif;
            line-height: 1.6;
            font-size: 11pt; /* Ensure font size is specified in points (pt) */
            margin: 0;
        }
        .container {
            max-width: calc(210mm - 32mm);
            margin: 0 auto;
            background: #fff;
            padding: 20px 16mm;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            position: relative; /* Ensure relative positioning for absolute positioning of logo */
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-top: 10px;
        }
        .logo {
            width: 150px;
            height: auto;
            opacity: 0.6;
            position: absolute;
            top: 10px; /* Adjust top position for desired vertical alignment */
            right: 20px; /* Adjust right position for right-hand corner */
        }
        .sidebar-images {
            position: absolute;
            left: 20px; /* Adjust left position for left-hand corner */
            top: 100px; /* Adjust top position for desired vertical alignment */
            display: flex;
            flex-direction: column;
            align-items: flex-start;
        }
        .sidebar-images img {
            width: 50px; /* Adjust width of images */
            height: auto; /* Maintain aspect ratio */
            margin-bottom: 10px; /* Adjust margin as needed */
        }
        .sidebar-images img.yellow {
            /* Specific styles for yellow image */
        }
        .sidebar-images img.black {
            /* Specific styles for black image */
        }
        h3 {
        font-size: 16pt; /* Specify font size in points (pt) */
        color: black;
        border-bottom: 1px solid #ccc;
        padding-bottom: 10px;
        margin-top: 20px;
        font-weight: normal; /* Ensures the text is not bold */
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background-color: #f0f0f0;
        }

        table th,
        table td {
            border: 1px solid #ffffff;
            text-align: left;
            padding: 1px;
            margin: 3px;
            margin-bottom: 10px;
            font-family: Calibri, Arial, sans-serif;
            vertical-align: middle; /* Vertically align text to the middle */
        }

        th {
            background-color: #e7e7e7;
            width: 25%;
            white-space: nowrap;
        }
        .profile {
            font-style: normal;
            margin-bottom: 20px;
        }
        .profile-section {
            display: flex;
            align-items: center;
            margin-bottom: 20px;
        }
        .profile-img {
            width: 120px;
            height: 120px;
            border-radius: 8px;
            margin-right: 20px;
        }
        .personal-details {
            flex: 1;
        }
        .bold-orange {
            font-weight: bold;
            color: rgb(252, 194, 6);
        }
    </style>
</head>

<body>

    <div class="container">
        ${html}
        <!-- Additional content can go here -->
    </div>
</body>
</html>
`;
    return template;
}

// Event listener for the download button
document.getElementById('downloadButton').addEventListener('click', exportToWord);

// File upload handling
const courseFile = document.getElementById('courseFile');
const dropZone = document.querySelector('.file-drop-zone');
const fileInfo = document.getElementById('file-info');
const fileName = fileInfo.querySelector('.file-name');

// Course file upload handling
document.getElementById('uploadCoursesBtn').addEventListener('click', function() {
    document.getElementById('courseFile').click();
});

courseFile.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    handleCourseUpload(file);
});

function handleCourseUpload(file, skipUpdate = true) {
    return new Promise((resolve, reject) => {
        if (!file.name.endsWith('.csv')) {
            reject('Please upload a CSV file');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        const btn = document.getElementById('uploadCoursesBtn');
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                if (!skipUpdate) {
                    updateFileInfo(file);
                }
                document.getElementById('coursesSection').classList.remove('hidden');
                document.getElementById('generateBtn').classList.remove('hidden');
                
                fetch('/?fetch_courses=true')
                    .then(response => response.json())
                    .then(data => {
                        updateCoursesTable(data.courses);
                        updateFilters(data.departments, data.semesters);
                    });
                resolve(data);
            } else {
                reject(data.error || 'Upload failed');
            }
        })
        .catch(reject)
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-upload mr-2"></i>Upload Course Data';
        });
    });
}

// Drag and drop handlers
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragging');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragging');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragging');
    const files = Array.from(e.dataTransfer.files);
    handleMultipleFiles(files);
});

// Update file info display
function updateFileInfo(file) {
    if (file) {
        fileName.textContent = file.name;
        fileInfo.classList.remove('hidden');
    } else {
        fileInfo.classList.add('hidden');
    }
}

// Room upload handling
function handleRoomUpload(file, skipUpdate = true) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);

        const btn = document.getElementById('uploadRoomsBtn');
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

        fetch('/upload-rooms', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                if (!skipUpdate) {
                    alert('Room data uploaded successfully');
                }
                resolve(data);
            } else {
                reject(data.error || 'Upload failed');
            }
        })
        .catch(reject)
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-building mr-2"></i>Upload Room Data';
        });
    });
}

// Batch upload handling
function handleBatchUpload(file, skipUpdate = true) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);

        const btn = document.getElementById('uploadBatchBtn');
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

        fetch('/upload-batches', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                if (!skipUpdate) {
                    alert('Batch data uploaded successfully');
                }
                resolve(data);
            } else {
                reject(data.error || 'Upload failed');
            }
        })
        .catch(reject)
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-users mr-2"></i>Upload Batch Data';
        });
    });
}

// Reserved slots upload handling  
function handleReservedUpload(file, skipUpdate = true) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);

        const btn = document.getElementById('uploadReservedBtn');
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

        fetch('/upload-reserved', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                if (!skipUpdate) {
                    alert('Reserved slots uploaded successfully');
                }
                resolve(data);
            } else {
                reject(data.error || 'Upload failed');
            }
        })
        .catch(reject)
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-clock mr-2"></i>Upload Reserved Slots';
        });
    });
}

// Faculty file upload handling
document.getElementById('uploadFacultyBtn').addEventListener('click', function() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv';
    input.onchange = e => {
        const file = e.target.files[0];
        if (file) {
            handleFacultyUpload(file, false);
        }
    };
    input.click();
});

function handleFacultyUpload(file, skipUpdate = true) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);

        const btn = document.getElementById('uploadFacultyBtn');
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

        fetch('/upload-faculty', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.text().then(text => {
                try {
                    return JSON.parse(text);
                } catch (e) {
                    throw new Error('Invalid JSON response from server');
                }
            });
        })
        .then(data => {
            if (data.success) {
                if (!skipUpdate) {
                    alert('Faculty data uploaded successfully');
                }
                resolve(data);
            } else {
                reject(data.error || 'Upload failed');
            }
        })
        .catch(error => {
            console.error('Upload error:', error);
            reject(error.message || 'Upload failed');
        })
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-user-tie mr-2"></i>Upload Faculty Data';
        });
    });
}

// Add click handler for reserved slots button
document.getElementById('uploadReservedBtn').addEventListener('click', function() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv';
    input.onchange = e => {
        const file = e.target.files[0];
        handleReservedUpload(file, false);
    };
    input.click();
});

// Add elective file upload handling
document.getElementById('uploadElectiveBtn').addEventListener('click', function() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv';
    input.onchange = e => {
        const file = e.target.files[0];
        if (file) {
            handleElectiveUpload(file, false);
        }
    };
    input.click();
});

function handleElectiveUpload(file, skipUpdate = true) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);

        const btn = document.getElementById('uploadElectiveBtn');
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

        fetch('/upload-elective-registrations', {  // Updated endpoint URL
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.text().then(text => {
                try {
                    return JSON.parse(text);
                } catch (e) {
                    throw new Error('Invalid JSON response from server');
                }
            });
        })
        .then(data => {
            if (data.success) {
                if (!skipUpdate) {
                    alert('Elective registrations uploaded successfully');
                }
                resolve(data);
            } else {
                reject(data.error || 'Upload failed');
            }
        })
        .catch(error => {
            console.error('Upload error:', error);
            reject(error.message || 'Upload failed');
        })
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-list-alt mr-2"></i>Upload Electives';
        });
    });
}

// Add generate button handler
document.getElementById('generateBtn').addEventListener('click', function(e) {
    e.preventDefault();
    
    const btn = this;
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Generating...';
    
    // Get duration values
    const durations = {
        lecture_duration: parseInt(document.getElementById('lectureDuration').value),
        lab_duration: parseInt(document.getElementById('labDuration').value),
        tutorial_duration: parseInt(document.getElementById('tutorialDuration').value),
        self_study_duration: 2,
        break_duration: 1,
        hour_slots: 2
    };

    // Save config then generate
    fetch('/save-config', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({duration_constants: durations})
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            document.getElementById('generateForm').submit();
        } else {
            throw new Error(data.error || 'Failed to save configuration');
        }
    })
    .catch(error => {
        alert('Failed to generate timetables: ' + error.message);
    })
    .finally(() => {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-cog mr-2"></i>Generate Timetables';
    });
});

// New function to update courses table
function updateCoursesTable(courses) {
    const tbody = coursesTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    courses.forEach(course => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-gray-50';
        tr.dataset.department = course.Department;
        tr.dataset.semester = course.Semester;
        
        tr.innerHTML = `
            <td class="border px-4 py-2">${course.Department}</td>
            <td class="border px-4 py-2">${course.Semester}</td>
            <td class="border px-4 py-2">${course['Course Code']}</td>
            <td class="border px-4 py-2">${course['Course Name']}</td>
            <td class="border px-4 py-2">${course.Faculty}</td>
            <td class="border px-4 py-2">${course.L}</td>
            <td class="border px-4 py-2">${course.T}</td>
            <td class="border px-4 py-2">${course.P}</td>
            <td class="border px-4 py-2">${course.Classroom}</td>
        `;
        tbody.appendChild(tr);
    });
}

// New function to update filters
function updateFilters(departments, semesters) {
    const deptFilter = document.getElementById('departmentFilter');
    const semFilter = document.getElementById('semesterFilter');
    
    // Update department filter
    deptFilter.innerHTML = '<option value="">All Departments</option>';
    departments.forEach(dept => {
        deptFilter.innerHTML += `<option value="${dept}">${dept}</option>`;
    });
    
    // Group semesters by their numeric part
    const semesterGroups = new Map();
    semesters.forEach(sem => {
        const semStr = String(sem);
        const numPart = semStr.match(/\d+/)[0];
        const sectionPart = semStr.match(/[A-Za-z]+/);
        
        if (!semesterGroups.has(numPart)) {
            semesterGroups.set(numPart, []);
        }
        if (sectionPart) {
            semesterGroups.get(numPart).push({
                value: sem,
                section: sectionPart[0]
            });
        }
    });
    
    // Update semester filter with grouped options
    semFilter.innerHTML = '<option value="">All Semesters</option>';
    
    // Add main semester numbers first
    Array.from(semesterGroups.keys()).sort((a, b) => parseInt(a) - parseInt(b)).forEach(num => {
        semFilter.innerHTML += `<option value="${num}">Semester ${num}</option>`;
        
        // Add sections if they exist
        const sections = semesterGroups.get(num);
        if (sections.length > 0) {
            sections.forEach(section => {
                semFilter.innerHTML += `<option value="${section.value}">&nbsp;&nbsp;&nbsp;Section ${section.section}</option>`;
            });
        }
    });
}

// Filter functionality
const departmentFilter = document.getElementById('departmentFilter');
const semesterFilter = document.getElementById('semesterFilter');
const coursesTable = document.getElementById('coursesTable');

function filterCourses() {
    const selectedDept = departmentFilter.value;
    const selectedSem = semesterFilter.value;
    
    const rows = coursesTable.querySelectorAll('tbody tr');
    
    rows.forEach(row => {
        const dept = row.dataset.department;
        const sem = row.dataset.semester;
        
        // For semester, match the numeric part if no section is specified
        const deptMatch = !selectedDept || dept === selectedDept;
        const semMatch = !selectedSem || 
            (selectedSem.match(/^\d+$/) ? 
                String(sem).startsWith(selectedSem) : // Match just the number
                String(sem) === String(selectedSem)); // Match exactly if section is specified
        
        row.classList.toggle('hidden', !(deptMatch && semMatch));
    });
}

departmentFilter.addEventListener('change', filterCourses);
semesterFilter.addEventListener('change', filterCourses);

// Bulk upload handling
document.getElementById('uploadAllBtn').addEventListener('click', function() {
    document.getElementById('allFiles').click();
});

document.getElementById('allFiles').addEventListener('change', function(e) {
    const files = Array.from(e.target.files);
    handleMultipleFiles(files);
});

function handleMultipleFiles(files) {
    const btn = document.getElementById('uploadAllBtn');
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Uploading...';

    // Validate all files are CSV
    if (!Array.from(files).every(file => file.name.endsWith('.csv'))) {
        alert('Please upload only CSV files');
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-upload mr-2"></i>Upload All Files';
        return;
    }

    // Process all files
    Promise.all(files.map(file => {
        if (file.name.toLowerCase().includes('room')) {
            return handleRoomUpload(file, false);
        } else if (file.name.toLowerCase().includes('batch')) {
            return handleBatchUpload(file, false);
        } else if (file.name.toLowerCase().includes('combined')) {
            return handleCourseUpload(file, false);
        } else if (file.name.toLowerCase().includes('reserved')) {
            return handleReservedUpload(file, false);
        } else if (file.name.toLowerCase().includes('faculty')) {
            return handleFacultyUpload(file, false);
        } else if (file.name.toLowerCase().includes('elective')) {
            return handleElectiveUpload(file, false);
        }
        return Promise.reject(`Unknown file type: ${file.name}`);
    }))
    .then(() => {
        // Update UI only after all files are processed
        document.getElementById('coursesSection').classList.remove('hidden');
        document.getElementById('generateBtn').classList.remove('hidden');
        fetch('/?fetch_courses=true')
            .then(response => response.json())
            .then(data => {
                updateCoursesTable(data.courses);
                updateFilters(data.departments, data.semesters);
            });
        alert('All files uploaded successfully');
    })
    .catch(error => {
        alert('Upload failed: ' + error);
    })
    .finally(() => {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-upload mr-2"></i>Upload All Files';
    });
}

// Duration Settings Modal
const durationModal = document.getElementById('durationModal');
const openDurationBtn = document.getElementById('openDurationSettings');
const saveDurationsBtn = document.getElementById('saveDurations');
const resetDurationsBtn = document.getElementById('resetDurations');

// Default durations
const defaultDurations = {
    lecture: 3,
    lab: 4,
    tutorial: 2
};

// Open modal
openDurationBtn.addEventListener('click', () => {
    durationModal.classList.remove('hidden');
});

// Close modal when clicking outside
durationModal.addEventListener('click', (e) => {
    if (e.target === durationModal) {
        durationModal.classList.add('hidden');
    }
});

// Update duration displays
function updateDurationDisplay(type, value) {
    const hours = (value * 30) / 60;
    const display = document.getElementById(`${type}DurationDisplay`);
    display.textContent = `${hours} hour${hours !== 1 ? 's' : ''} (${value} slots)`;
}

// Add input listeners for sliders
['lecture', 'lab', 'tutorial'].forEach(type => {
    const slider = document.getElementById(`${type}Duration`);
    slider.addEventListener('input', () => {
        updateDurationDisplay(type, slider.value);
    });
});

// Reset durations
resetDurationsBtn.addEventListener('click', () => {
    Object.entries(defaultDurations).forEach(([type, value]) => {
        const slider = document.getElementById(`${type}Duration`);
        slider.value = value;
        updateDurationDisplay(type, value);
    });
});

// Save durations
saveDurationsBtn.addEventListener('click', () => {
    const durations = {
        lecture_duration: parseInt(document.getElementById('lectureDuration').value),
        lab_duration: parseInt(document.getElementById('labDuration').value),
        tutorial_duration: parseInt(document.getElementById('tutorialDuration').value),
        self_study_duration: 2, // Keep default
        break_duration: 1, // Keep default
        hour_slots: 2 // Keep constant
    };

    fetch('/save-config', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({duration_constants: durations})
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            durationModal.classList.add('hidden');
        }
    });
});

// Initialize displays
['lecture', 'lab', 'tutorial'].forEach(type => {
    const slider = document.getElementById(`${type}Duration`);
    updateDurationDisplay(type, slider.value);
});

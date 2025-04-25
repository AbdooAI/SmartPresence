%%     CJ    %%
%% HASHEM VASEGHI 22004087
%% ABDELRAHMAN KHALAFALLA 22019464

%% Load the data
clear;
clc;
s = serialport("COM7",9600,'Timeout', 1209600); % serial port between matlab and arduino
fileName = 'Attendence.xlsx'; % Define the name of the Excel file

%%
% Define course data with student IDs, names, and RFID tags
courses.CMP101.STD =  [];
courses.CMP101.name = [];
courses.CMP101.Rfid = [];
% load data from excel:
data_CMP101 = readtable('students.xlsx', 'Sheet', "CMP101");
% Convert table columns to cell arrays
newSTD = table2cell(data_CMP101(:, 'STD'));
newName = table2cell(data_CMP101(:, 'name'));
newRfid = table2cell(data_CMP101(:, 'Rfid'));
% Append new data to the existing structure
courses.CMP101.STD = [courses.CMP101.STD, newSTD'];
courses.CMP101.name = [courses.CMP101.name, newName'];
courses.CMP101.Rfid = [courses.CMP101.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.MATH201.STD =  [];
courses.MATH201.name = [];
courses.MATH201.Rfid = [];
% load data from excel:
data_MATH201 = readtable('students.xlsx', 'Sheet', "MATH201");
% Convert table columns to cell arrays
newSTD = table2cell(data_MATH201(:, 'STD'));
newName = table2cell(data_MATH201(:, 'name'));
newRfid = table2cell(data_MATH201(:, 'Rfid'));
% Append new data to the existing structure
courses.MATH201.STD = [courses.MATH201.STD, newSTD'];
courses.MATH201.name = [courses.MATH201.name, newName'];
courses.MATH201.Rfid = [courses.MATH201.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.BIO102.STD =  [];
courses.BIO102.name = [];
courses.BIO102.Rfid = [];
% load data from excel:
data_BIO102 = readtable('students.xlsx', 'Sheet', "BIO102");
% Convert table columns to cell arrays
newSTD = table2cell(data_BIO102(:, 'STD'));
newName = table2cell(data_BIO102(:, 'name'));
newRfid = table2cell(data_BIO102(:, 'Rfid'));
% Append new data to the existing structure
courses.BIO102.STD = [courses.BIO102.STD, newSTD'];
courses.BIO102.name = [courses.BIO102.name, newName'];
courses.BIO102.Rfid = [courses.BIO102.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.ENG202.STD =  [];
courses.ENG202.name = [];
courses.ENG202.Rfid = [];
% load data from excel:
data_ENG202 = readtable('students.xlsx', 'Sheet', "ENG202");
% Convert table columns to cell arrays
newSTD = table2cell(data_ENG202(:, 'STD'));
newName = table2cell(data_ENG202(:, 'name'));
newRfid = table2cell(data_ENG202(:, 'Rfid'));
% Append new data to the existing structure
courses.ENG202.STD = [courses.ENG202.STD, newSTD'];
courses.ENG202.name = [courses.ENG202.name, newName'];
courses.ENG202.Rfid = [courses.ENG202.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.PHY203.STD =  [];
courses.PHY203.name = [];
courses.PHY203.Rfid = [];
% load data from excel:
data_PHY203 = readtable('students.xlsx', 'Sheet', "PHY203");
% Convert table columns to cell arrays
newSTD = table2cell(data_PHY203(:, 'STD'));
newName = table2cell(data_PHY203(:, 'name'));
newRfid = table2cell(data_PHY203(:, 'Rfid'));
% Append new data to the existing structure
courses.PHY203.STD = [courses.PHY203.STD, newSTD'];
courses.PHY203.name = [courses.PHY203.name, newName'];
courses.PHY203.Rfid = [courses.PHY203.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.HIST301.STD =  [];
courses.HIST301.name = [];
courses.HIST301.Rfid = [];
% load data from excel:
data_HIST301 = readtable('students.xlsx', 'Sheet', "HIST301");
% Convert table columns to cell arrays
newSTD = table2cell(data_HIST301(:, 'STD'));
newName = table2cell(data_HIST301(:, 'name'));
newRfid = table2cell(data_HIST301(:, 'Rfid'));
% Append new data to the existing structure
courses.HIST301.STD = [courses.HIST301.STD, newSTD'];
courses.HIST301.name = [courses.HIST301.name, newName'];
courses.HIST301.Rfid = [courses.HIST301.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.PSY302.STD =  [];
courses.PSY302.name = [];
courses.PSY302.Rfid = [];
% load data from excel:
data_PSY302 = readtable('students.xlsx', 'Sheet', "PSY302");
% Convert table columns to cell arrays
newSTD = table2cell(data_PSY302(:, 'STD'));
newName = table2cell(data_PSY302(:, 'name'));
newRfid = table2cell(data_PSY302(:, 'Rfid'));
% Append new data to the existing structure
courses.PSY302.STD = [courses.PSY302.STD, newSTD'];
courses.PSY302.name = [courses.PSY302.name, newName'];
courses.PSY302.Rfid = [courses.PSY302.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.ART304.STD =  [];
courses.ART304.name = [];
courses.ART304.Rfid = [];
% load data from excel:
data_ART304 = readtable('students.xlsx', 'Sheet', "ART304");
% Convert table columns to cell arrays
newSTD = table2cell(data_ART304(:, 'STD'));
newName = table2cell(data_ART304(:, 'name'));
newRfid = table2cell(data_ART304(:, 'Rfid'));
% Append new data to the existing structure
courses.ART304.STD = [courses.ART304.STD, newSTD'];
courses.ART304.name = [courses.ART304.name, newName'];
courses.ART304.Rfid = [courses.ART304.Rfid, newRfid'];
%%

% Define course data with student IDs, names, and RFID tags
courses.CHEM204.STD =  [];
courses.CHEM204.name = [];
courses.CHEM204.Rfid = [];
% load data from excel:
data_CHEM204 = readtable('students.xlsx', 'Sheet', "CHEM204");
% Convert table columns to cell arrays
newSTD = table2cell(data_CHEM204(:, 'STD'));
newName = table2cell(data_CHEM204(:, 'name'));
newRfid = table2cell(data_CHEM204(:, 'Rfid'));
% Append new data to the existing structure
courses.CHEM204.STD = [courses.CHEM204.STD, newSTD'];
courses.CHEM204.name = [courses.CHEM204.name, newName'];
courses.CHEM204.Rfid = [courses.CHEM204.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.STAT403.STD =  [];
courses.STAT403.name = [];
courses.STAT403.Rfid = [];
% load data from excel:
data_STAT403 = readtable('students.xlsx', 'Sheet', "STAT403");
% Convert table columns to cell arrays
newSTD = table2cell(data_STAT403(:, 'STD'));
newName = table2cell(data_STAT403(:, 'name'));
newRfid = table2cell(data_STAT403(:, 'Rfid'));
% Append new data to the existing structure
courses.STAT403.STD = [courses.STAT403.STD, newSTD'];
courses.STAT403.name = [courses.STAT403.name, newName'];
courses.STAT403.Rfid = [courses.STAT403.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.SOC303.STD =  [];
courses.SOC303.name = [];
courses.SOC303.Rfid = [];
% load data from excel:
data_SOC303 = readtable('students.xlsx', 'Sheet', "SOC303");
% Convert table columns to cell arrays
newSTD = table2cell(data_SOC303(:, 'STD'));
newName = table2cell(data_SOC303(:, 'name'));
newRfid = table2cell(data_SOC303(:, 'Rfid'));
% Append new data to the existing structure
courses.SOC303.STD = [courses.SOC303.STD, newSTD'];
courses.SOC303.name = [courses.SOC303.name, newName'];
courses.SOC303.Rfid = [courses.SOC303.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.PHIL404.STD =  [];
courses.PHIL404.name = [];
courses.PHIL404.Rfid = [];
% load data from excel:
data_PHIL404 = readtable('students.xlsx', 'Sheet', "PHIL404");
% Convert table columns to cell arrays
newSTD = table2cell(data_PHIL404(:, 'STD'));
newName = table2cell(data_PHIL404(:, 'name'));
newRfid = table2cell(data_PHIL404(:, 'Rfid'));
% Append new data to the existing structure
courses.PHIL404.STD = [courses.PHIL404.STD, newSTD'];
courses.PHIL404.name = [courses.PHIL404.name, newName'];
courses.PHIL404.Rfid = [courses.PHIL404.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.ECON401.STD =  [];
courses.ECON401.name = [];
courses.ECON401.Rfid = [];
% load data from excel:
data_ECON401 = readtable('students.xlsx', 'Sheet', "ECON401");
% Convert table columns to cell arrays
newSTD = table2cell(data_ECON401(:, 'STD'));
newName = table2cell(data_ECON401(:, 'name'));
newRfid = table2cell(data_ECON401(:, 'Rfid'));
% Append new data to the existing structure
courses.ECON401.STD = [courses.ECON401.STD, newSTD'];
courses.ECON401.name = [courses.ECON401.name, newName'];
courses.ECON401.Rfid = [courses.ECON401.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.POLS402.STD =  [];
courses.POLS402.name = [];
courses.POLS402.Rfid = [];
% load data from excel:
data_POLS402 = readtable('students.xlsx', 'Sheet', "POLS402");
% Convert table columns to cell arrays
newSTD = table2cell(data_POLS402(:, 'STD'));
newName = table2cell(data_POLS402(:, 'name'));
newRfid = table2cell(data_POLS402(:, 'Rfid'));
% Append new data to the existing structure
courses.POLS402.STD = [courses.POLS402.STD, newSTD'];
courses.POLS402.name = [courses.POLS402.name, newName'];
courses.POLS402.Rfid = [courses.POLS402.Rfid, newRfid'];
%%
% Define course data with student IDs, names, and RFID tags
courses.COMP501.STD =  [];
courses.COMP501.name = [];
courses.COMP501.Rfid = [];
% load data from excel:
data_COMP501 = readtable('students.xlsx', 'Sheet', "COMP501");
% Convert table columns to cell arrays
newSTD = table2cell(data_COMP501(:, 'STD'));
newName = table2cell(data_COMP501(:, 'name'));
newRfid = table2cell(data_COMP501(:, 'Rfid'));
% Append new data to the existing structure
courses.COMP501.STD = [courses.COMP501.STD, newSTD'];
courses.COMP501.name = [courses.COMP501.name, newName'];
courses.COMP501.Rfid = [courses.COMP501.Rfid, newRfid'];
%% 

% Define course data with student IDs, names, and RFID tags
courses.ELEE342.STD =  [];
courses.ELEE342.name = [];
courses.ELEE342.Rfid = [];
% load data from excel:
data_ELEE342 = readtable('students.xlsx', 'Sheet', "ELEE342");
% Convert table columns to cell arrays
newSTD = table2cell(data_ELEE342(:, 'STD'));
newName = table2cell(data_ELEE342(:, 'name'));
newRfid = table2cell(data_ELEE342(:, 'Rfid'));
% Append new data to the existing structure
courses.ELEE342.STD = [courses.ELEE342.STD, newSTD'];
courses.ELEE342.name = [courses.ELEE342.name, newName'];
courses.ELEE342.Rfid = [courses.ELEE342.Rfid, newRfid'];


%%
courses.ELEE464.STD = [];
courses.ELEE464.name = [];
courses.ELEE464.Rfid = [];
data_ELEE464 = readtable('students.xlsx', 'Sheet', "ELEE464");
% Convert table columns to cell arrays
newSTD = table2cell(data_ELEE464(:, 'STD'));
newName = table2cell(data_ELEE464(:, 'name'));
newRfid = table2cell(data_ELEE464(:, 'Rfid'));
% Append new data to the existing structure
courses.ELEE464.STD = [courses.ELEE464.STD, newSTD'];
courses.ELEE464.name = [courses.ELEE464.name, newName'];
courses.ELEE464.Rfid = [courses.ELEE464.Rfid, newRfid'];

%%
courses.ELEE306.STD = [];
courses.ELEE306.name = [];
courses.ELEE306.Rfid = [];

data_ELEE306 = readtable('students.xlsx', 'Sheet', "ELEE306");
% Convert table columns to cell arrays
newSTD = table2cell(data_ELEE306(:, 'STD'));
newName = table2cell(data_ELEE306(:, 'name'));
newRfid = table2cell(data_ELEE306(:, 'Rfid'));
% Append new data to the existing structure
courses.ELEE306.STD = [courses.ELEE306.STD, newSTD'];
courses.ELEE306.name = [courses.ELEE306.name, newName'];
courses.ELEE306.Rfid = [courses.ELEE306.Rfid, newRfid'];
%%


% Get the field names (course names)
courseNames = fieldnames(courses);

% Loop through each course and write to the Excel sheet
for i = 1:numel(courseNames)
    courseName = courseNames{i};
    % Extract the STD and name fields
    data.STD = courses.(courseName).STD';
    data.name = courses.(courseName).name';
    
    % Convert to table
    T = struct2table(data);
    
    % Write to Excel with the course name as the sheet name
    writetable(T, fileName, 'Sheet', courseName);
end

while 1
    % RFID check
    providedRfid = readline(s); % Read RFID input from the serial port
    providedRfid = strtrim(regexprep(providedRfid, '\n', '')); % Clean up the RFID input

    % Check time
    currentDateTime = datetime('now'); % Get current date and time
    dayOfWeek = strjoin(day(currentDateTime, 'shortname')); % Get the day of the week
    currentTime = datestr(currentDateTime, 'HH:MM'); % Get current time as a string
    dateHeader = sprintf('%d/%d/%d', str2double(datestr(now, 'mm')), ...
        str2double(datestr(now,'dd')), str2double(datestr(now,'yyyy'))); % Current date in 'MM/DD/YYYY' format

    % Extract hours and minutes
    currentHour = hour(currentDateTime);
    currentMinute = minute(currentDateTime);

    if currentMinute < 30
        adjustedHour = currentHour - 1;
    else
        adjustedHour = currentHour;
    end
    adjustedMinute = 30;
    adjustedTimeStr = sprintf('%02d:%02d', adjustedHour, adjustedMinute); % Adjusted time string

    % adjustedTimeStr = '9:30'; % Fixed time for testing
    % dayOfWeek = 'Tue'; % Fixed day for testing
    % dateHeader='6/9/2024'; % Fixed date for testing
    % % Define the structure to store class schedule data
    classSchedule = struct(...
        'days', {{"Sun",'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'}}, ...
        'hours', {{'8:30', '9:30', '10:30', '11:30', '12:30', '13:30', '14:30', '15:30', '16:30', '17:30'}}, ...
        'courses', {{...
            %   sun   ,   Mon   ,   Tue   ,   Wed   ,   Thu   ,   Fri   ,   Sat   
            {   []    ,'CMP101' ,   []    ,'CHEM204',   []    ,'ELEE306',   []    }, ... % 8:30
            {   []    ,'CMP101' ,'PHY203' ,'CHEM204',   []    ,'ELEE306',   []    }, ... % 9:30
            {   []    ,'MATH201','PHY203' ,'CHEM204','ECON401',"ELEE342",   []    }, ... % 10:30
            {   []    ,'MATH201',   []    ,   []    ,'ECON401',"ELEE342",   []    }, ... % 11:30
            {   []    ,   []    ,'HIST301','STAT403','ECON401','ELEE464',   []    }, ... % 12:30
            {   []    ,'BIO102' ,'HIST301','STAT403',   []    ,'ELEE464',   []    }, ... % 13:30
            {   []    ,'BIO102' ,'PSY302' ,'SOC303' ,'POLS402',   []    ,   []    }, ... % 14:30
            {   []    ,'ENG202' ,'PSY302' ,   []    ,'COMP501',   []    ,   []    }, ... % 15:30
            {   []    ,'ENG202' ,   []    ,'PHIL404','COMP501',   []    ,   []    }, ... % 16:30
            {   []    ,   []    ,'ART304' ,'PHIL404','COMP501',   []    ,   []    } ... % 17:30
        }} ...
    );
   
    % Find the index of the specified day
    dayIndex = find(strcmp(classSchedule.days, dayOfWeek));

    % Find the index of the specified time
    timeIndex = find(strcmp(classSchedule.hours, adjustedTimeStr));

    % Get the course scheduled at the specified day and time
    if ~isempty(dayIndex) && ~isempty(timeIndex)
        course = classSchedule.courses{timeIndex}{dayIndex};
    else
        course = 'No course found';
        continue; % Skip to the next iteration of the loop
    end
 if isempty(course)

       continue; % Skip to the next iteration of the loop 
    end
    % Access the course data dynamically if a course is found
    if isfield(courses, course)
        courseData = courses.(course);
        
        % Find the index of the provided RFID tag in the course data
        rfidIndex = find(strcmp(courseData.Rfid, providedRfid));

        if ~isempty(rfidIndex)
            studentID = courseData.STD(rfidIndex);
            studentName = courseData.name(rfidIndex);
            disp(['RFID Tag belongs to Student ID: ', studentID]);
            disp(['Student Name: ', studentName]);
            Attendence = true;
        else
            disp('RFID Tag not found in the course data');
            continue; % Skip to the next iteration of the loop
        end
    else
        disp('Course data not available');
    end

    % Read and write to Excel sheet
    sheetName = course;

    % Check if the file exists
    if isfile(fileName)
        % Read existing headers
        [~, ~, raw] = xlsread(fileName, sheetName, 'A1:Z1');
        existingHeaders = raw(1, :);
    else
        % If the file doesn't exist, set existingHeaders to an empty cell array
        warning('File does not exist, headers not found.');
        existingHeaders = {};
    end

    indexfound = false;
    is_nan_block = cellfun(@(x) isnumeric(x) && all(isnan(x)), existingHeaders);
    indexfound = false;
    i = 1;

    while indexfound == false
        if strcmp(existingHeaders{i}, dateHeader)
            indexfound = true;
        elseif is_nan_block(i) == 1
            indexfound = true;
        else
            i = i + 1;
        end
    end

    existingHeaders{i} = dateHeader; % Update headers with current date
    xlswrite(fileName, existingHeaders, sheetName, 'A1'); % Write updated headers to Excel

    columnLetter = char('A' + (i - 1)); % Calculate the column letter for the date
    cellReference = sprintf('%s%d', columnLetter, (rfidIndex + 1)); % Determine the cell reference for the student

    % Write "Present" to the specified cell
    xlswrite(fileName, {"Present"}, sheetName, cellReference);
end

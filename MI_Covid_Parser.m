tic
% experimenting with reading directly from the website:
%fname = "https://www.michigan.gov/coronavirus/-/media/Project/Websites/coronavirus/Michigan-Data/11-22-2022/Cases-and-Deaths-by-County-and-by-Date-of-Symptom-Onset-or-by-Date-of-Death2022-11-22.xlsx"
%urlwrite(fname,'test.xlsx');
% stores spreadsheet to test.xlsx.
% need to handle datesformat1 = 'mm-dd-yyyy';
% datestr(now,format1);

% fname = "https://www.michigan.gov/coronavirus/-/media/Project/Websites/coronavirus/Michigan-Data/";
% fdate = "11-29-2022";
% fname = strcat(fname,fdate);
% fname2 = "/Cases-and-Deaths-by-County-and-by-Date-of-Symptom-Onset-or-by-Date-of-Death";
% fname = strcat(fname,fname2);
% fdate2 = "2022-11-29.xlsx";
% fname = strcat(fname,fdate2);

% Setup the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 8);

% Specify sheet and range
opts.Sheet = "Data";
opts.DataRange = "A2";

% Specify column names and types
opts.VariableNames = ["COUNTY", "Date", "CASE_STATUS", "Cases", "Deaths", "CasesCumulative", "DeathsCumulative", "Updated"];
opts.VariableTypes = ["categorical", "datetime", "categorical", "double", "double", "double", "double", "datetime"];

% Specify variable properties
opts = setvaropts(opts, ["COUNTY", "CASE_STATUS"], "EmptyFieldRule", "auto");
opts = setvaropts(opts, "Date", "InputFormat", "");
opts = setvaropts(opts, "Updated", "InputFormat", "");

% Import the data
[fname, path] = uigetfile('*.xlsx','Select Cases and Deaths File');
oldPath = cd(path);
tbl = readtable(fname, opts, "UseExcel", false);

%% Convert to output type
COUNTY = tbl.COUNTY;
Date = tbl.Date;
CASE_STATUS = tbl.CASE_STATUS;
Cases = tbl.Cases;
Deaths = tbl.Deaths;
CasesCumulative = tbl.CasesCumulative;
DeathsCumulative = tbl.DeathsCumulative;
Updated = tbl.Updated;

% Clear temporary variables
clear opts

%%
% Get total of all counties for each date (confirmed only)
d = unique(Date);
tmp = ~isnat(d);
day = d(tmp);
days = size(day,1);
for i = 1:days
    d = day(i);
    daySlice = tbl(Date == d & CASE_STATUS == 'Confirmed',:);
    case_sum(i) = sum(daySlice.Cases);
    death_sum(i) = sum(daySlice.Deaths);
end

% Output file
date_string = datestr(Updated(2),'mm-dd-yy');
file_str = sprintf('covid_data_%s.txt',date_string);
Fid = fopen(file_str,'w');
fprintf('Michigan       Total Cases: %6d Total Deaths: %6d Death Pct: %5.3f%%\n',sum(case_sum),sum(death_sum),100.0*sum(death_sum)/sum(case_sum));
fprintf(Fid,'Michigan       Total Cases: %6d Total Deaths: %6d Death Pct: %5.3f%%\n',sum(case_sum),sum(death_sum),100.0*sum(death_sum)/sum(case_sum));
% Apply seven day smoothing to total state data
windowSize = 7;
b = (1/windowSize)*ones(1,windowSize);
a = 1;
smoothed_cases = filter(b,a,case_sum);
smoothed_deaths = filter(b,a,death_sum);

% Set last two days to NaN because they may be incomplete
smoothed_cases(end-1:end) = NaN;
smoothed_deaths(end-1:end) = NaN;

figure;
bar(day,case_sum);
hold on;
plot(day,smoothed_cases,'m .-');
title('Michigan Cases vs date');
legend('Raw Data','Seven Day Average');
xlabel('Date');
ylabel('Cases');
grid on;
f = gcf;
file_str = sprintf('MI_Cases_%s.png',date_string);
exportgraphics(f,file_str);

figure;
bar(day,death_sum);
hold on;
plot(day,smoothed_deaths,'m .-');
title('Michigan Deaths vs date');
legend('Raw Data','Seven Day Average');
xlabel('Date');
ylabel('Deaths');
grid on;
f = gcf;
file_str = sprintf('MI_Deaths_%s.png',date_string);
exportgraphics(f,file_str);

%%
% Get specific county data
counties = ["Ottawa","Newaygo"];
county_strings = ["Ottawa ",...
                  "Newaygo"];
for c = 1:2
    total_cases(c) = 0;
    total_deaths(c) = 0;
    county = counties(c);
    countySlice = tbl(COUNTY == county & CASE_STATUS == 'Confirmed' & ~isnat(Date),:);
    countySlice = sortrows(countySlice,'Date');
    county_cases = countySlice.Cases;
    county_deaths = countySlice.Deaths;
    total_cases(c) = sum(county_cases);
    total_deaths(c) = sum(county_deaths);

    % Apply seven day smoothing to county data
    windowSize = 7;
    b = (1/windowSize)*ones(1,windowSize);
    a = 1;
    smoothed_county_cases = filter(b,a,county_cases);
    smoothed_county_deaths = filter(b,a,county_deaths);

    % Set last two days to NaN because they may be incomplete
    smoothed_county_cases(end-1:end) = NaN;
    smoothed_county_deaths(end-1:end) = NaN;

    figure;
    bar(day,county_cases);
    hold on;
    plot(day,smoothed_county_cases,'m .-');
    title([county,' County Cases vs date']);
    legend('Raw Data','Seven Day Average');
    xlabel('Date');
    ylabel('Cases');
    grid on;
    f = gcf;
    file_str = sprintf('%s_County_Cases_%s.png',county,date_string);
    exportgraphics(f,file_str);

    figure;
    bar(day,county_deaths);
    hold on;
    plot(day,smoothed_county_deaths,'m .-');
    title([county,' County Deaths vs date']);
    legend('Raw Data','Seven Day Average');
    xlabel('Date');
    ylabel('Deaths');
    grid on;
    f = gcf;
    file_str = sprintf('%s_County_Deaths_%s.png',county,date_string);
    exportgraphics(f,file_str);
    
    clear county_cases county_deaths
    fprintf('%s County Total Cases: %6d Total Deaths: %6d Death Pct: %5.3f%%\n',county_strings(c),total_cases(c),total_deaths(c),100.0*total_deaths(c)/total_cases(c));
    fprintf(Fid,'%s County Total Cases: %6d Total Deaths: %6d Death Pct: %5.3f%%\n',county_strings(c),total_cases(c),total_deaths(c),100.0*total_deaths(c)/total_cases(c));
    
end    
fclose(Fid);
toc
cd(oldPath);

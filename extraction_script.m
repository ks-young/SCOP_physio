% Extracting EMG data - SCOP project

%TO DO
% set up additional output data (error output?)
% annotate

clear all; close all; clc
%set up data locations
main_dir = 'C:\Users\kyoung\google Drive\other_work\SCOP_programming';
data_dir = strcat(main_dir,'\test-files');
cd(data_dir)

%find data files & print out list
file_list = dir('00*');
fprintf('Compiling data file list: \n')
for a = 1:length(file_list)
    file_name{a,1} = file_list(a).name;
    fprintf(strcat(file_name{a,1},'\n'))
end

error_count = 0;
%%
% EMG extraction
for b = 1:length(file_list)
    fprintf(strcat('Reading data file: \t', file_name{b}, '...\n'))
    audio = xlsread(file_name{b,1},'Parameters','E:E');
    emg = xlsread(file_name{b,1},'Parameters','G:G');
    [data txt raw] = xlsread(file_name{b,1},'Parameters','N:O');
    sub(b).fname = file_name{b};
    
    fprintf('Finding startle markers \n')
    txt = txt(3:length(txt),:);
    
    type = {'BASELINE STARTLE', 'ANTICIPATION STARTLE'...
            'SPEECH STARTLE', 'ITI'};
    code = {'base','sp_ant','sp_startle','sp_ITI'};
    trials = {'12';'14';'14';'21'};
        
    for bx= 1:length(type)
        index = strfind(txt(:,2),type{bx});
        onsets = find(~cellfun(@isempty,index));
        sub(b).type(bx) = code(bx);
        
        fprintf(strcat('Extracting data type: ',type{bx},'\n'))
        for c = 1:length(onsets)
            %find peak audio output
            audio_range = audio(onsets(c):onsets(c)+199,1);
            peak_audio = max(audio_range);
            
            if peak_audio < 100
                fprintf(strcat('checking for pre-marker peak in trial: ',num2str(c),'\n'))
                audio_range2 = audio(onsets(c)-100:onsets(c),1);
                peak_audio = max(audio_range2);
                index_peak = onsets(c)-100+find(audio_range2==peak_audio)-1;
                error_count = error_count + 1;                
                if peak_audio < 100
                    fprintf('no valid startle\n')
                    sub(b).emg_error(c,bx) = {'no startle audio'};
                elseif peak_audio >= 100
                    sub(b).emg_error(c,bx) = {'startle before marker'};
                    sub(b).emg_error_idx(c,bx) = -(onsets(c) - index_peak);
                end
            else fprintf('peak found\n')
                index_peak = onsets(c)+find(audio_range==peak_audio);
            end
            
            %find peak EMG
            emg_max_range = emg(index_peak+1:index_peak+15,1);
            peak_emg = max(emg_max_range);
            sub(b).emg_max(c,bx) = peak_emg;
            
            %find baseline EMG
            emg_base_range = emg(index_peak-21:index_peak-2,1);
            base_emg = mean(emg_base_range);
            sub(b).emg_bl(c,bx) = base_emg;
        end
    end
    clear data txt raw audio emg
    fprintf(strcat('---- finished extraction for file: \t', file_name{b}, '\n'));
end
%%
% reformat and output data
out_data{1,1}= 'PID';
for aa=1:12 %12 baseline startles
    out_data{1,aa+1} = strcat('BASE_',num2str(aa));
end

for ab=1:7 %7 speeches
    for ac=1:2 %2 anticipation startles/speech
        aa=aa+1;
        out_data{1,aa+1} = strcat('SP_ANT_',num2str(ab),'.',num2str(ac));
    end
end

for ab=1:7 %7 speeches
    for ac=1:2 %2 startles/speech
        aa=aa+1;
        out_data{1,aa+1} = strcat('SP_STARTLE_',num2str(ab),'.',num2str(ac));
    end
end

for ab=1:7 %7 speeches
    for ac=1:3 %3 ITI startles/speech
        aa=aa+1;
        out_data{1,aa+1} = strcat('SP_ITI_',num2str(ab),'.',num2str(ac));
    end
end

for b=1:length(sub)
    PID_code = char(file_name(b)); %update depending on file naming system
    PID = PID_code(1:4);
    out_data{b+1,1} = PID;
    max_data{b+1,1} = PID;
    baseline_data{b+1,1} = PID;
    col = 1;
    for c=1:length(trials) %4 types of trials
        for d=1:str2num(trials{c})
            max = sub(b).emg_max(d,c);
            baseline = sub(b).emg_bl(d,c);
            emg_resp = max - baseline;
            out_data{b+1,col+1} = emg_resp;
            max_data{b+1,col+1} = max;
            baseline_data{b+1,col+1} = baseline;
            col=col+1;
        end
    end
end

    fprintf('Writing output file...');
     filename = 'emg_extracted.xlsx';
             xlswrite(filename,out_data,'emg_resp','A1')
             xlswrite(filename,max_data,'emg_max','A1')
             xlswrite(filename,out_data(1,:),'emg_max','A1')
             xlswrite(filename,baseline_data,'emg_baseline','A1')
             xlswrite(filename,out_data(1,:),'emg_baseline','A1')
                  
printf('SCRIPT FINISHED-----END OF PROCESSING-----\n')
%% This script opens a NetCDF array, retreives the groups imbeded
%% and creates excel files for each of them. Within each files are dimensions
%% representing time, depth, variables, etc. which are then tabs of the excel files.'

clear
fprintf('MARS Netcdf unwrapper \n');

%% Dictionnaries
models =    {'GFDL-ESM2M';...
    %'HadGEM2-ES';...
    'IPSL-CM5A-LR'};...
    %'MIROC-ESM-CHEM';...
%'NorESM1-M'};

stations   = { 'My_station',                59.75, 7.25;...
 };

%% Climate model loop . only 1 and 3 are selected for Task 4.4 of MARS project

for mod_loop = 1:2
    
    current_model = models(mod_loop);
    %% Change to match local folder structure 
    Working_folder = strcat('D:\...\MARS\CLIMATE\',current_model,'\');
    
    %% Station loop
    
    
    
    for stn_loop = 1:length(stations(:,1))
        
        station_name = stations(stn_loop,1);
        
        %% Open dir list in a txt file made in advance, create an array with dir names
        Dirlist_file = 'dirlist.txt';
        formatSpec = '%12s%[^\n\r]';
        fileID = fopen(Dirlist_file,'r');
        dataArray = textscan(fileID, formatSpec, 'Delimiter', '', 'WhiteSpace', '',  'ReturnOnError', false);
        fclose(fileID);
        dir_list = dataArray{:, 1};
        clearvars filename formatSpec fileID dataArray ans;
        
        fprintf ((strcat('Parsing NetCDF_',current_model{1},' for station_',station_name{1},'\n')));
        
        %% Excel API
        
        xlsx_file = strjoin(strcat('D:\iLandDrive\MARS\CLIMATE\MARS_climate_',current_model,'_',station_name,'.xlsx'));
        xlswrite(xlsx_file,'A1')
        %xlswrite(FILE,ARRAY,SHEET,RANGE)
        Excel = actxserver ('Excel.Application');
        File=xlsx_file;
        
        if ~exist(File,'file')
            ExcelWorkbook = Excel.workbooks.Add;
            ExcelWorkbook.SaveAs(File,1);
            ExcelWorkbook.Close(false);
        end
        
        invoke(Excel.Workbooks,'Open',File);
        
        
        %%  loop through directory lists
        for folder_i = 1:numel(dir_list)
            
            run_folder = dir_list(folder_i); % moving to folder
            PathName = strjoin(strcat(Working_folder,run_folder)); % 1st concatentes, then create a string with strjoin
            cd(PathName)
            % var name is a function of run_folder minus the path info rcp8p5\
            temp = run_folder{1};
            var_name = temp(8:end); % the ffirst 8 char are gargage
            
            %% file name loop by year
            yr_column = 0; % to write columns in sequence in the output matrix
            
            for yr_loop=2006:2099
                
                yr_column = yr_column + 1;
                % filename must change based in folder name ...
                
                if str2double(dir_list{folder_i}(4)) == 4 %% making sure to have the right filename
                    FileName = strjoin(strcat(var_name,'_bced_1960_1999_',current_model,'_rcp4p5_',num2str(yr_loop),'.nc'));
                elseif str2double(dir_list{folder_i}(4)) == 8
                    FileName = strjoin(strcat(var_name,'_bced_1960_1999_',current_model,'_rcp8p5_',num2str(yr_loop),'.nc'));
                end
                
                ncid = netcdf.open(FileName, 'NC_NOWRITE');
                [ndims,nvars,ngatts,unlimdimid] = netcdf.inq(ncid); %get information
                
                for k = 0:nvars-1
                    [varname,xtype,dimids,natts] = netcdf.inqVar(ncid,k); % get vars and theyr type in group
                    %fprintf('Var %d is %s  (type = %d) \n', k, varname, xtype)
                    
                    for j = 0:ndims-1  % dimsnsions : 0-time, 1-lat, 2-lon
                        [dimname, dimlen] = netcdf.inqDim(ncid,j); %get dims and their lenght in group
                        %fprintf('Dim %d is %s (lenght = %d) \n', j, dimname, dimlen)
                    end
                    
                end
                
                time_vec = netcdf.getVar(ncid,0);
                
                % Serial date correction for MATLAB datum to EXCEL
                time_vec = time_vec + 679352 - 693960; % the file starts at 1860-1-1
                
                lat_vec = netcdf.getVar(ncid,1);
                lon_vec = netcdf.getVar(ncid,2);
                var_vec = netcdf.getVar(ncid,5);
                
                %% Interpolation routine to from 0.5 to 0.1 degree
                % var_vec should be interpolated ...
                
                %% retrive specific latxlon for all times set to your catchment
                
                lat_ind = find(lat_vec==stations{stn_loop,2});
                lon_ind = find(lon_vec==stations{stn_loop,3}); %
                temp = var_vec(lon_ind,lat_ind,:);
                % concatenation
                if yr_column == 1
                    output = horzcat(time_vec, temp(1,:)');
                else
                    data_block = horzcat(time_vec, temp(1,:)'); %
                    output = vertcat(output, data_block);
                end
                
                netcdf.close(ncid)

                
            end
            
            % write in excel, each folder is a tab
            temp = strjoin(run_folder); % making a sheet name string
            sheet = strrep(temp, '\','-'); % replacing the \ with -
            
            %xlswrite(FILE,ARRAY,SHEET,RANGE)
            xlswrite1(File,output,sheet,'A1')
            
            
        end
        
        invoke(Excel.ActiveWorkbook,'Save');
        Excel.Quit
        Excel.delete
        clear Excel
        
        fclose all
        
    end
end

cd ../../../


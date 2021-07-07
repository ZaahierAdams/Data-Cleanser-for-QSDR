'''

Contains all text for Help menu 

'''

def Changelog_text():
    Changes = '''
[29/02/2020]
 
 -  Provided exceptions for a number of possible errors
         - Missing Persal & TI DBs
         - Wrong sheet names for DBs
         - Wrong column names in DBs
         - Missing pdf docs
         - Missing system images
 - Most of these exceptions will produce a notification in Feedback window
_________________________________________________________________________________________________
[27/02/2020]
 
 -  Added version information to executable file
 -  Lightened Persal DB
_________________________________________________________________________________________________
[26/02/2020]
 
 -  Built in exceptions for checks:
         -  Residual columns verifiers
         -  TI verifiers
         -  'Format_Excel_Cell' method

 -  Disabled TI text outputs
_________________________________________________________________________________________________
[25/02/2020]
 
 -  Changed unverified TI cells highlight from #ff9563 to #fff45c 
_________________________________________________________________________________________________
[24/02/2020]
 
 -  Built verifiers for 17 columns that cannot be verified by any baseline data
         -  Checks for validity of the column values
         -  Strength of verifications range from weak to robust
         -  All 17 columns represented as a single variable in results outputs
_________________________________________________________________________________________________
[23/02/2020]
 
 -  Output QMR now preserves value for 'Disability' column, where it differs to Persal
 -  Given unique, pink, cell highlight
 
 -  Now highlights all erroneous and unverified Persal cells (as opposed to ID only)
 -  Highlights all unverified TI cells 
_________________________________________________________________________________________________
[22/02/2020]
 
 -  Training Intervention (TI) DB now incorporated
 -  Used to verify Training Interventions in QMR
 -  Built in script that saves TI's recognised and unrecognised by TIDB
 -  Each TI issue represented in output graph and Result Reports 
_________________________________________________________________________________________________
[21/02/2020]
 
 -  Verifies three additional columns from Persal:
         - Residence
         - Occupation Level
         - Job Title
 -  New persal checks represented in output graph and Result Reports 
 -  tkinter graphics now opens as maximized window 
_________________________________________________________________________________________________
[14/02/2020]
 
 -  Updated VBA macro
_________________________________________________________________________________________________
 [13/02/2020]
 
 -  Renamed Output directories:
         'VBA' -> 'Formatted Reports'  
         'XLSX_simple' -> 'Simple Reports'
         'Sheets' -> 'Separate Sheets'
         
 -  Renamed 'Output Format' Radio Buttons:
         'VBA' -> 'Formatted Report'
         'Simple XLSX' -> 'Simple Report'
         'Separate' -> 'Separate Sheets'
         
 -  Updated User Guide (edition 1.2)
_________________________________________________________________________________________________
 [11/02/2020]
 
 -  Added script for button tooltips (inactive - commented out)
 -  Updated User Guide (edition 1.1)

_________________________________________________________________________________________________
 [10/02/2020]
 
 -  Added text 'Result reports' notification in feedback window
 -  Added 'input' and 'output' dir. buttons to gui 
 -  'Graphics' and 'Output Format' slave windows have new .ico

 -  Renamed dir. 'Dirt_Reports' to 'Result Reports'  
 -  Renamed dir. 'Test_Excel_Files' to 'Input files'
 -  Renamed dir. 'New_Excel_Files' to 'Output files'
 -  Renamed dir. 'Result_Diagrams' to 'Result Diagrams'
 -  Changed .ico for 'Find file' button
 -  'Graphics' and 'Output Format' slave windows open at (dx,dy) to root window
 -  'Cleanse Data' button hover highlight changed
 -  After cleansing, Feedback window now jumps to last notification 
_________________________________________________________________________________________________
 [09/02/2020] 
 
 -  Matplotlib problem in .exe fixed by downgrading Numpy (from 1.18.0? to to 1.15.0) 




'''
    return Changes


def Description():
    Description ='''
About this application:

• Detects and corrects all 
  errors in QSDR excel 
  workbooks
  
• Outputs 'cleaned' QSDR 
  excel workbooks

• Provides details of some of 
  the issues resolved by 
  means of graphics and excel 
  reports
  

_______________________

Developed by:
 Zaahier Adams
 github.com/ZaahierAdams

'''
    return Description


def Version(version_string, last_updated):
    
    Version = (  'Data Cleanser for QSDR workbooks\n'
               + 'Version {}\n\n'.format(version_string)
               + 'Created on:\t20/01/2020\n'
               + 'Lasted updated:\t{}'.format(last_updated)
               )
    

    return Version

def Report():
    Report = '''
Please communicate to us with us via email:
Email (developer): Zaahier Adams



Please structure the email as follows:
(i) Title    Stating the nature of the issue: 
                • Question
                • Bug
                • Help wanted
                • etc...
                
(ii) Description



If it is a Bug that you are reporting:
(1) Please describe what happened,
(2) Preferably attach a screenshot of any error message 
(3) Additionally, attach the excel file(s) that you attempted to clean
'''

    return Report



    
























def Ascii_Art():
    Ascii = '''


                                 (%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  
                         (%%%%%                                             %%  
                    /%%%                                                    %%  
                 %%%                                                        %%  
              %%%                                                           %%  
           .%%                      *                           %           %%  
         .%%      /%*        %     .%% %% %%*     (           #%%%          %%  
        %%     %%%%//.,       %.(%*%%%% (#%%% %% %          ./%%%%%*        %%  
      *%(     %%*%%%.% %%      . %   %%%%%%.  %           % %%%%%           %%  
     %%       %%%.%%*%%%%%      %%  %#.  .(%  %(          ( %%%%%           %%  
    #%         %%%%%%%%(   %                        %%%  %%%%%%%            %%  
    %*          (.%%%%%%  /%    %%     ,%  #%%#%%%  %  %%%%% %%/            %%  
   %%      #%,   %%%%%,    %     (     %% %% %%% %% % *%,  /%%%%   %*       %%  
   %,       %  %%%%%      %% ,   (    .%%%  %%%%%   %%   *  .%%%%  %        %%  
  #%       %  %%%%*  (  #%#%   %%%%   %%%%.  ((%    % ,%    ,  %%%, %       %%  
  %%        %%%%%% %    .  %         %%%%%%,        %        ,%%% */        %%  
  (%          %%%  #/  .    (      %%%%%%%%%%             %%*%%%%           %/  
   %#         %  %,% (       .    %%%      %%%.    .     %%% /%%%%         #%   
   %%        # #      %       %.%%%          %%%((      ,%      % %        %%   
    %%        /%          %    /%%%%        (%%%%       *       #%        %%    
     %(        , * %%  (%%%%%%%%%*%%%%    %%%%##%%%%%%%%%  %%            (%     
      %%        %%%%%%%#   %%%%%%%% %%%%%%%% %%%%%%%.   /%%%%%%%,       %%      
       %%       %  ,%%#,(    .%%%%%%%%% .*%%%%%%%%     %(*%%/  %       %%       
        #%#          %%   #.%%   /(           %   %##%   #%          (%#        
          %%*           %%% */ ,%(            /%/ ,( %%%           *%%          
            %%%                   %%(.     (%%.                  %%%            
               %%(                                            (%%               
                  %%%                                      %%%                  
                      %%%%                            %%%%        
                      

'''
    return Ascii







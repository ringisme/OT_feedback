# %reset -f

import os
import pandas as pd
from glob import glob


# Set working direction:
os.chdir("F:\\Workshop\\01_Match8_feedback")

# Detect and import Excel:
file_name = glob('Source\\*.xls')
raw_df = [pd.read_excel(f) for f in file_name][0]

# Rename columns:
raw_df.columns = ['ID', 'cmt_date', 'response_time', 'source',
                  'source_details', 'IP', 'Q1_teaching_date', 'Q2_kindergarten',
                  'Q3_class_type', 'Q4_1', 'Q4_2', 'Q5_1',
                  'Q5_2', 'Q6_appear_on_time', 'Q7_Preparation', 'Q8_Enthusiasm',
                  'Q9_Interesting', 'Q10_fluency', 'Q11_student_feedback', 'Q12_arrive_late',
                  'Q13_late_time', 'Q14_overall_comments', 'Q15_late_reason', 'Q16_online_teacher',
                  'Q17_class_teacher', 'total_score']

# Create ref teachers tables:
ref_file_name = glob('OT_ref\\*.csv')
OT_ref = pd.read_csv(ref_file_name[0])

# Clean the data initially:
ext_df = raw_df[['Q16_online_teacher', 'Q6_appear_on_time', 'Q7_Preparation', 
        'Q8_Enthusiasm', 'Q9_Interesting', 'Q10_fluency', 
        'Q11_student_feedback','total_score', 'Q14_overall_comments'
        ]]

ext_df.columns = ['OTID', 'appear_on_time', 'preparation', 
                  'enthusiasm', 'interaction', 'fluency', 
                  'student_feedback', 'total_score', 'comments']

ext_df = pd.merge(ext_df, OT_ref, on='OTID', how = 'left')
ext_feedback = pd.concat([ext_df['OT_name'], ext_df.iloc[:, 1:9]], axis =1)

# Create Pivot table:
def group_feedback(df):
    groupMean = round(df.mean(), 2)
    if groupMean > 5:
        return str(groupMean) + ' of 30'
    else:
        return str(groupMean) + ' of 5'

clean_feedback_p1 = ext_feedback.groupby('OT_name').agg(group_feedback)
clean_feedback_p2 = ext_feedback.groupby('OT_name').apply(lambda x: len(x.index))
clean_feedback = pd.concat([clean_feedback_p1, clean_feedback_p2], axis = 1)

clean_feedback = clean_feedback.rename(columns = {0:'total_class'})

# For comments:
def group_comment(text):
    seperator = ' || '
    return seperator.join(text)

raw_comment = ext_feedback.groupby('OT_name').agg(group_comment)

trans_comment = raw_comment.copy()

for i in range(0, raw_comment.size):
    print(raw_comment.iloc[i, :])
    trans_input = input('Please transfer the comment:')
    trans_comment.iloc[i, 0] = trans_input

final_feedback = pd.merge(clean_feedback, trans_comment, on='OT_name', how='left')

AskForOutput = input('Output as xlsx? (y/n)')
if AskForOutput == 'y':
    final_feedback.to_excel("Oct28_to_Nov3_feedback.xlsx", sheet_name='feedback')
    print('Excel file has been created.')
else:
    print('cancelled.')

to_end = input('press to end.')




















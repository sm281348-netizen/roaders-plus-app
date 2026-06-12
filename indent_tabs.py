import re

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_target_tab = False
tab_names_to_hide = ['tab1', 'tab_m', 'tab6', 'tab3', 'tab4', 'tab5', 'tab7']
target_tabs_regex = re.compile(r'^with (' + '|'.join(tab_names_to_hide) + r'):')

for i, line in enumerate(lines):
    # Check if we are starting a target tab block
    match = target_tabs_regex.match(line)
    if match:
        new_lines.append(f'if current_hotel != "採購":\n')
        new_lines.append('    ' + line)
        in_target_tab = True
        continue
    
    # If we are inside a target tab block, we need to indent
    if in_target_tab:
        # Check if the block has ended (a line with no indentation, except empty lines)
        if line.strip() != '' and not line.startswith(' ') and not line.startswith('\t'):
            in_target_tab = False
            new_lines.append(line)
        else:
            # Indent by 4 spaces
            if line == '\n':
                new_lines.append(line)
            else:
                new_lines.append('    ' + line)
    else:
        new_lines.append(line)

with open('app.py', 'w', encoding='utf-8', newline='\n') as f:
    f.writelines(new_lines)

print("Done indenting tabs!")

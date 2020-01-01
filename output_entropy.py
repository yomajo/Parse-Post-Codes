import os
import random

# GLOBAL VARIABLES
OUTPUT_FILENAME = 'post_codes.txt'
codes_count = 5000
code_num_len = 9
formatting_len = code_num_len
original_example_string = 'UA015125015LTUA015125024LTUA015125038LTUA015125041LT - 9 nums'

def gen_rand(dec_nums):
    ''''generates random number [0;dec_nums]'''
    return random.randint(0, 10**dec_nums)

def gen_list(n_elem):
    '''outputs a list of unique n_elem elements'''
    code_list = []
    if n_elem <= 10**code_num_len:
        for _ in range(n_elem):
            temp_code = gen_rand(code_num_len)
            while temp_code in code_list:
                temp_code = gen_rand(code_num_len)
            code_list.append(temp_code)
            # print(f'Adding unique {temp_code}')
    else:
        raise Exception(f'You have asked for more codes THAN random generator could generate', f'codes count: {n_elem}; generator capacity: {10**code_num_len}')
    return code_list

def format_litems(some_list, dec_places):
    '''takes list as arg and converts each member to string with leading zeros in total of dec_places str len'''
    for idx, value in enumerate(some_list):
        some_list[idx] = 'UA' + str(value).zfill(dec_places) + 'LT'
    return some_list

def list_to_str(some_list):
    '''converts some_list to continuous string'''
    return ''.join(some_list)

def output_txt(string_output):
    '''outputs passed string to a text file'''
    file_dir = os.path.dirname(__file__)
    output_dir = os.path.join(file_dir, OUTPUT_FILENAME)
    with open(output_dir, 'w') as f:
        f.write(string_output)
    print(f'Check output at: {output_dir}')
    return output_dir

def run():
    codes = gen_list(codes_count)
    codes_mod = format_litems(codes, formatting_len)
    string_output = list_to_str(codes_mod)
    output_txt(string_output)
        
if __name__ == "__main__":
    run()
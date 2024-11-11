import subprocess
import sys
import json
import os

# Make sure the user download the package
def package_download(package):
    try:
        __import__(package)
    except ImportError:
        print(f"Package '{package} not found. Installing it now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

package_download("pandas")
package_download("openpyxl")
import pandas as pd
from openpyxl import Workbook

def main():
    file_name = input("유저 리스트 파일명 (확장자 제외) 입력: ")
    user_data = []
    json_file_path = os.path.join(os.path.dirname(os.path.relpath(__file__)), f"{file_name}.json")
    excel_file_path = os.path.join(os.path.dirname(os.path.relpath(__file__)), f"{file_name}.xlsx")

    # Check if JSON file exists
    if not os.path.exists(json_file_path):
        user_input = input("유저 리스트가 없습니다. 엑셀 파일에서 불러 오시겠습니까? Y/N: ").lower()
        
        if user_input == 'y':
            if os.path.exists(excel_file_path):
                user_data = load_data_from_excel(excel_file_path)
                print(f"총 {len(user_data)}명의 유저 데이터가 불러와졌습니다.")
            else:
                print(f"엑셀 파일이 존재하지 않습니다")
                user_data = load_or_create_json(json_file_path)
    else:
        user_data = load_or_create_json(json_file_path)

    # main
    while True:
        print("\n유저 리스트 수정: 1\n유저 리스트 보기: 2\n공대 파티 계산: 3\n그만 두기: 4")
        action = int(input())
        match action:
            case 1:
                modify_character(user_data)
            case 2:
                display_characters(user_data)
            case 3:
                party_size = int(input("\n공대 파티 인원수를 입력: "))
                parties = create_party(user_data, party_size)
                print_party(parties) 
            case 4:
                save_json(json_file_path, user_data)
                sys.exit()
            case _:
                print("\n잘못된 입력입니다")

def modify_character(user_data):
    display_characters(user_data)
    while True:
        add_user = input("\n새로운 유저를 추가하시겠습니까? Y/N: ").lower()
        if add_user == 'y':
            username = input("유저 이름: ").strip()
            new_user = {'user_name': username, 'characters': []}
            user_data.append(new_user)
            print(f"유저 '{username}' 추가 완료")

        elif add_user == 'n':
            username = input("\n편집하고 싶은 유저 이름: (그만 두려면 'exit', 유저를 삭제하려면 'remove', 현재 리스트를 보려면 'look'): ").strip().lower()
            
            if username.lower() == 'exit':
                break
            elif username.lower() == 'remove':
                username_to_remove = input("\n삭제하려는 유저 이름: ").strip()
                user_data = [user for user in user_data if user['user_name'] != username_to_remove]
                print(f"유저 '{username_to_remove}' 삭제 완료.")
            else:
                user_found = False
                for user in user_data:
                    if user['user_name'] == username:
                        user_found = True
                        print(f"{username}: ")
                        for idx, char in enumerate(user['characters']):
                            print(f"  {idx + 1}. {char['character_name']} - 직업: {char['character_class']}, 템렙: {char['character_level']}")

                        char_idx = int(input("\n변경하려는 캐릭터 번호 입력: ")) - 1
                        if 0 <= char_idx < len(user['characters']):
                            char = user['characters'][char_idx]
                            print(f"변경하려는 캐릭터: {char['character_name']}")
                            new_class = input(f"직업 변경 (현재: {char['character_class']}): ")
                            new_level = input(f"템렙 변경 (현재: {char['character_level']}): ")

                            if new_class:
                                char['level'] = int(new_class)
                            if new_level:
                                char['character_level'] = int(new_level)

                            print(f"캐릭터 {char['character_name']} 변경 완료")
                        else:
                            print("올바르지 않은 캐릭터 번호.")
                            add_new = input("새 캐릭터를 추가 하시겠습니까? Y/N: ").lower()
                            if add_new == 'y': add_new_character(user)
                            else: break
                        break

                if not user_found:
                    print(f"유저 '{username}' 리스트에 없음.")

        else: 
            print("잘못된 입력입니다")

def add_new_character(user):
    # print(f"\n'{user['username']}'의 새 캐릭터 추가: ")
    name = input("이름 입력: ").strip()
    char_class = input("직업 입력: ").strip()
    level = int(input("템렙 입력: "))
    
    new_character = {
        "character_name": name,
        "character_class": char_class,
        "character_level": level,
    }
    
    user['characters'].append(new_character)
    print(f"캐릭터 '{name}' 추가 완료")

def display_characters(user_data):
    if not user_data:
        print("\n빈 리스트 입니다")
    else:
        for i, user in enumerate(user_data):
            print(f"\n{'='*40}")
            print(f"유저 {i + 1}: {user['user_name']}")
            print(f"{'-'*40}")
            
            if user['characters']:
                for char_i, char in enumerate(user['characters']):
                    print(f"\n캐릭터 {char_i + 1}:")
                    print(f"  이름      : {char['character_name']}")
                    print(f"  직업      : {char['character_class']}")
                    print(f"  템렙      : {char['character_level']}")
                print(f"{'='*40}")
            else:
                print("  캐릭터가 없습니다.")
            

def create_party(user_data, party_size):
    parties = []
    
    all_characters = []
    for user in user_data:
        username = user["user_name"]
        for char in user["characters"]:
            all_characters.append({"user_name": username, 
            "character_name": char["character_name"], 
            "character_class": char["character_class"], 
            "character_level": char["character_level"]})
    
    all_characters.sort(key=lambda x: x["character_level"], reverse=True)
    
    while len(all_characters) >= party_size:
        party = []
        for _ in range(party_size):
            party.append(all_characters.pop(0)) 
        parties.append(party)
    
    if all_characters:
        parties.append(all_characters)

    balance_parties(parties)
    
    return parties

def balance_parties(parties):
    party_levels = [sum(char["character_level"] for char in party) for party in parties]

    if len(parties) == 0: 
        return
    
    while True:
        avg_level = sum(party_levels) / len(party_levels)
        
        highest_party_idx = party_levels.index(max(party_levels))
        lowest_party_idx = party_levels.index(min(party_levels))
        
        if party_levels[highest_party_idx] - party_levels[lowest_party_idx] > avg_level:
            highest_party = parties[highest_party_idx]
            lowest_party = parties[lowest_party_idx]
            
            highest_char = max(highest_party, key=lambda x: x["character_level"])
            lowest_char = min(lowest_party, key=lambda x: x["character_level"])
            
            highest_party.remove(highest_char)
            lowest_party.remove(lowest_char)
            
            highest_party.append(lowest_char)
            lowest_party.append(highest_char)
            
            party_levels[highest_party_idx] = sum(char["character_level"] for char in highest_party)
            party_levels[lowest_party_idx] = sum(char["character_level"] for char in lowest_party)
        else:
            break  

import pandas as pd

def print_party(parties):
    if len(parties) == 0: 
        print("현재 유저 리스트가 비어 있습니다")
        return

    total_level = 0  
    char_num = 0
    party_data = []  
    
    for i, party in enumerate(parties):
        print(f"\n{'='*40}")
        print(f"공대 {i + 1}:")
        print(f"{'-'*40}")
        
        party_level = 0  
        party_characters = []  
        
        for character in party:
            print(f"\n캐릭터명 : {character['character_name']}")
            print(f"  직업     : {character['character_class']}")
            print(f"  템렙     : {character['character_level']}")
            
            party_characters.append({
                "공대 번호": i + 1,
                "캐릭터명": character['character_name'],
                "직업": character['character_class'],
                "템렙": character['character_level']
            })
            
            party_level += character["character_level"]
            char_num += 1
        
        party_average = party_level / len(party)
        total_level += party_level
        print(f"\n파티 템렙 평균: {party_average:.2f}")
        print(f"{'='*40}")
        
        party_data.extend(party_characters)

    average_level = total_level / char_num if char_num > 0 else 0
    print(f"\n총 템렙 평균: {average_level:.2f}")
    
    export_choice = input("\n결과를 Excel 파일로 내보내시겠습니까? (Y/N): ").strip().lower()
    if export_choice == 'y':
        df = pd.DataFrame(party_data)
        excel_file = "공대 파티 데이터.xlsx"
        df.to_excel(excel_file, index=False)
        print(f"Excel 파일 '{excel_file}'로 내보내기 완료!")
    else:
        print("Excel 파일로 내보내지 않았습니다.")

def load_or_create_json(file_path):
    if not os.path.exists(file_path):
        print(f"{file_path}가 없습니다. 새 리스트 만드는 중...")
        return [] 
    else:
        with open(file_path, "r", encoding="utf-8") as file:
            return json.load(file)

def save_json(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(data, file, indent=4, ensure_ascii=False)  
    print(f"리스트가 {file_path}에 저장 되었습니다")

def load_data_from_excel(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path, engine="openpyxl")
    
    # Ensure that the Excel file has the correct columns
    required_columns = ['유저 이름', '캐릭터 이름', '직업', '템렙']
    if not all(col in df.columns for col in required_columns):
        print(f"Excel 파일에 필요한 열이 없습니다. 필요한 열: {required_columns}")
        return []
    
    user_data = []
    # Process the Excel data
    for _, row in df.iterrows():
        user_name = row['유저 이름']
        character_name = row['캐릭터 이름']
        character_class = row['직업']
        character_level = int(row['템렙'])
        
        # Check if the user already exists
        user = next((user for user in user_data if user['user_name'] == user_name), None)
        if user is None:
            user = {'user_name': user_name, 'characters': []}
            user_data.append(user)
        
        # Add the character to the user's list
        user['characters'].append({
            'character_name': character_name,
            'character_class': character_class,
            'character_level': character_level
        })
    
    return user_data


if __name__ == "__main__":
    main()
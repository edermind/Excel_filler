import pandas as pd
import textdistance as tx


# Процент совпадения
def percent_coin(first: str, second: str) -> float:
    
    # Преобразуем строки в нижний регистр для корректного сравнения
    dist = tx.damerau_levenshtein.distance(first, second)
    max_len = max(len(first), len(second))
    percent = (1 - dist / max_len) * 100

    return percent


# Чтение таблиц
analyze_df = pd.read_excel('/home/buda/analyze.xlsx')
list_owner_df = pd.read_excel('/home/buda/list_owner.xlsx')


analyze_df['Процент_совпадения'] = pd.Series(dtype='float')
analyze_df['AD_Блок'] = pd.Series(dtype='str')

# Обработка данных
for index, row in analyze_df.iterrows():
    analyze_block: str = 'Н/Д'
    try:
        analyze_block = row['AD_Destination'].split('/')[3]
    except IndexError:
        analyze_df.at[index, 'AD_Блок'] = analyze_block
        analyze_df.at[index, 'Процент_совпадения'] = 0.0
        continue
    except Exception as e:
        continue
    
    bestOwner = ""
    bestCandidate = "Н/Д"
    bestPercent = 0.0
    for _, owner_row in list_owner_df.iterrows():
        percent = percent_coin(analyze_block.strip().lower(), owner_row['Владелец'].strip().lower())
        if percent > bestPercent:
            bestOwner = owner_row['Владелец']
            bestPercent = percent
            bestCandidate = owner_row['Блок']
    
    if bestPercent > 25.0:
        analyze_df.at[index, 'AD_Блок'] = bestCandidate
        analyze_df.at[index, 'Процент_совпадения'] = bestPercent
    else:
        analyze_df.at[index, 'AD_Блок'] = "Н/Д"
        analyze_df.at[index, 'Процент_совпадения'] = bestPercent
        


analyze_df.to_excel('answer.xlsx', index=False)

print("Скрипт выполнен. Результаты сохранены в 'answer.xlsx'")

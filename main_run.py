import pandas as pd
import path_finder
import dart
import to_excel
import kis_chrome
import post_office
import os
os.chdir('')


# df는 kischrome의 result로 통일 하였으며
# 조회되거나 검색하는 이름은 모두 종목명(업무배분_clean) 으로 통일


ss_path = ''
out_dir = ''
target_dir = '\\은행연합회'
posted_from = out_dir
posted_to = '\\평가내역'

# 담당마다 돌려
list1 = ['사람0', '사람1', '사람2', '사람3']
for damdang in list1:
    upmoo_damdang = damdang
    up_bae_total_dir = '{}.xlsx'.format(damdang)

    # 솔직히 이건 필요 없는데
    kischrome0 = kis_chrome.KisChrome(up_bae_total_dir)
    result = kischrome0.upmoo_df

    # path finder
    final_dirs = []
    logs = []
    for r in result.loc[:, '종목명(업무배분_clean)']:
        pathfinder0 = path_finder.PathFinder(out_dir, target_dir)
        final_dir = pathfinder0.match_and_save(r)
        final_dirs.append(final_dir)
        logs.append([final_dir.iloc[0, 0], pathfinder0.matched_folder,
                    pathfinder0.second_matched_folder])
        print(final_dir)

        # 종목 하나 만들었으면 포장해서 날려

        try:
            po0 = post_office.PostOffice(
                r, up_bae_total_dir, upmoo_damdang, posted_from, posted_to)
            # po0.copy_for_upmoo()
            po0.post_and_clear()
        except:
            print('error sending {}'.format(r))

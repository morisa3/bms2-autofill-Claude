import unittest

from bridgefill.site.session import (LOGGED_IN, LOGIN_PAGE, OTHER, UNKNOWN, classify, clean_user_agent, describe,
                                     parse_probe, redact_url)


class SessionTests(unittest.TestCase):
    def test_user_agent_drops_qtwebengine_but_keeps_chrome_version(self):
        agent = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                 'QtWebEngine/6.10.0 Chrome/134.0.6998.208 Safari/537.36')
        cleaned = clean_user_agent(agent)
        self.assertNotIn('QtWebEngine', cleaned)
        self.assertIn('Chrome/134.0.6998.208', cleaned)
        self.assertNotIn('  ', cleaned)

    def test_redact_url_removes_query_and_fragment(self):
        self.assertEqual(redact_url('https://bms2.iot.gov.tw/bms3/inspect/148156/?token=abc#top'),
                         'https://bms2.iot.gov.tw/bms3/inspect/148156/')
        self.assertEqual(redact_url(''), '')

    def test_classify_states(self):
        self.assertEqual(classify({'hasMenu': True, 'hasLoginForm': False}), LOGGED_IN)
        self.assertEqual(classify({'hasMenu': False, 'hasLoginForm': True}), LOGIN_PAGE)
        self.assertEqual(classify({'hasMenu': True, 'hasLoginForm': True}), LOGIN_PAGE)
        self.assertEqual(classify({'hasMenu': False, 'hasLoginForm': False}), OTHER)
        self.assertEqual(classify({}), UNKNOWN)

    def test_parse_probe_rejects_bad_values(self):
        self.assertEqual(parse_probe('{"hasMenu": true}'), {'hasMenu': True})
        self.assertEqual(parse_probe('not json'), {})
        self.assertEqual(parse_probe(None), {})
        self.assertEqual(parse_probe('[1, 2]'), {})

    def test_describe_lists_captcha_and_menu_without_query(self):
        login = describe({'url': 'https://x/login/?a=1', 'hasLoginForm': True, 'hasRecaptcha': True,
                          'hasImageCaptcha': True})
        self.assertEqual(login, '登入頁｜https://x/login/｜驗證：reCAPTCHA、圖片驗證碼')
        menu = describe({'url': 'https://x/home/', 'hasMenu': True, 'menuItems': ['基本資料', '檢測資料']})
        self.assertEqual(menu, '已登入｜https://x/home/｜選單 2 項：基本資料、檢測資料')


if __name__ == '__main__':
    unittest.main()

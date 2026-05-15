# inject_ga.py

import streamlit.components.v1 as components

GA_MEASUREMENT_ID = "G-6E63DFP24Z"  # 実際のGA4測定IDに置き換え

def inject_ga():
    ga_code = f"""
    <!-- Google tag (gtag.js) -->
    <script async src="https://www.googletagmanager.com/gtag/js?id={GA_MEASUREMENT_ID}"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){{dataLayer.push(arguments);}}
      gtag('js', new Date());
      gtag('config', '{GA_MEASUREMENT_ID}');
    </script>
    """

    components.html(ga_code, height=0, width=0)

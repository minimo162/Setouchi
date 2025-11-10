use leptos::*;

#[component]
pub fn AppShell() -> impl IntoView {
    view! {
        <main class="setouchi-app">
            <header class="hero">
                <p class="eyebrow">"Aurora Runway"</p>
                <h1>"Setouchi Excel Copilot"</h1>
                <p class="tagline">"セルから銀河まで翻訳の軌道をつなぐ。"</p>
            </header>
            <section class="command-dock">
                <h2>"フォーム (モック)"</h2>
                <div class="form-grid">
                    <label>"ブック"
                        <input placeholder="選択してください" />
                    </label>
                    <label>"シート"
                        <input placeholder="選択してください" />
                    </label>
                </div>
            </section>
            <section class="timeline">
                <h2>"Nebula Timeline"
                </h2>
                <ul>
                    <li>"まだメッセージはありません"</li>
                </ul>
            </section>
        </main>
    }
}

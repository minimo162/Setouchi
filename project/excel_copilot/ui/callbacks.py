# excel_copilot/ui/callbacks.py

import streamlit as st

def handle_user_input():
    """
    ユーザーがチャット入力でメッセージを送信したときに呼び出されるコールバック。
    """
    # ユーザーの入力を取得し、メッセージ履歴に追加
    if user_input := st.session_state.user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})

        # エージェントの処理を開始
        agent = st.session_state.agent
        excel_manager = st.session_state.excel_manager
        
        # 応答をストリーミングで受け取るための準備
        with st.chat_message("assistant"):
            # st.statusを使用してエージェントの思考プロセスを表示
            with st.status("エージェントが思考中...", expanded=True) as status:
                final_response_parts = []
                # agent.runはジェネレータ
                for chunk in agent.run(user_input, excel_manager):
                    if "最終回答" in chunk or "Final Answer" in chunk:
                        # 最終回答の前のログはステータスに表示
                        status.update(label="最終回答を生成中...", state="complete", expanded=False)
                    elif "思考:" in chunk or "アクション:" in chunk or "観察:" in chunk:
                         # 途中の思考プロセスはログとして表示
                         status.write(chunk)
                    else:
                        final_response_parts.append(chunk)

            # 最終的な回答をチャットに表示
            final_response = "".join(final_response_parts)
            st.write(final_response)
            st.session_state.messages.append({"role": "assistant", "content": final_response})
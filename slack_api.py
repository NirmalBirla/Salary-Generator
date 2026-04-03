from slack_sdk.errors import SlackApiError
import time

from connection import channel_id, client


def fetch_all_users():
    # Fetch user list
    users_list = {}
    try:
        result = client.users_list()
        for user in result["members"]:
            user_id = user["id"]
            real_name = user["profile"].get("real_name", "Unknown")
            email = user["profile"].get("email", "N/A")
            users_list[user_id] = {"real_name": real_name, "email": email}
    except SlackApiError as e:
        print(f"Error fetching user list: {e.response['error']}")
    return users_list

def retrieve_messages_with_threads(sleep_between_calls=1.0):
    """
    Returns a list of messages where each thread is bound to its parent message.
    - Only fetches threads when reply_count > 0
    - Clean nested structure: parent['replies'] = list of replies
    """
    all_messages = []

    try:
        cursor = None
        while True:
            # Fetch parent messages
            result = client.conversations_history(
                channel=channel_id,
                limit=1000,
                cursor=cursor
            )

            messages = result.get("messages", [])

            for msg in messages:
                # Add a copy of the parent message
                parent = msg.copy()

                # If this message has replies (a thread), fetch and bind them
                if msg.get("reply_count", 0) > 0:
                    thread_ts = msg.get("thread_ts") or msg.get("ts")
                    thread_messages = fetch_full_thread(client, channel_id, thread_ts, sleep_between_calls)

                    # replies list = all thread messages except the parent itself
                    replies = [tmsg for tmsg in thread_messages if tmsg.get("ts") != msg.get("ts")]
                    if replies:
                        parent["replies"] = replies

                all_messages.append(parent)

            # Handle pagination
            cursor = result.get("response_metadata", {}).get("next_cursor")
            if not cursor:
                break

            time.sleep(sleep_between_calls)

    except SlackApiError as e:
        print(f"Error fetching conversation history: {e.response['error']}")
        if e.response.get('error') == "ratelimited":
            print("Rate limited — try increasing sleep_between_calls")

    print(f"Total parent messages retrieved: {len(all_messages)}")
    return all_messages


def fetch_full_thread(client, channel_id, thread_ts, sleep=1.0):
    """Fetch complete thread (parent + all replies)"""
    thread_messages = []
    cursor = None

    try:
        while True:
            result = client.conversations_replies(
                channel=channel_id,
                ts=thread_ts,
                limit=1000,
                cursor=cursor
            )

            thread_messages.extend(result.get("messages", []))

            cursor = result.get("response_metadata", {}).get("next_cursor")
            if not cursor:
                break

            time.sleep(sleep)

    except SlackApiError as e:
        print(f"Error fetching thread {thread_ts}: {e.response['error']}")

    return thread_messages
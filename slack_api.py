from slack_sdk.errors import SlackApiError

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


def retrieve_messages():
    # Retrieve messages from the Slack channel
    try:
        limit = 1000
        result = client.conversations_history(channel=channel_id, limit=limit)
        return result["messages"]
    except SlackApiError as e:
        print(f"Error fetching conversation history: {e.response['error']}")
        return []
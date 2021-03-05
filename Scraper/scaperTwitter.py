from twitterscraper import query_tweets


# All tweets matching either Trump or Clinton will be returned. You will get at
# least 10 results within the minimal possible time/number of requests
for tweet in query_tweets("Trump", 10)[:10]:
    print(tweet.user.encode('utf-8'))


print("djfhskldghfeiggsffd")
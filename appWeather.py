import asyncio, os
from dotenv import load_dotenv
from air import DistillerClient

load_dotenv()
API_KEY = os.getenv("API_KEY")

PROJECT = "weather_project"   # letters/numbers/hyphens/underscores only
USER_ID = "test_user"         # same naming rule

async def main():
    # 1) Create/register the project once (uploads your YAML)
    client = DistillerClient(api_key=API_KEY)
    client.create_project(config_path="configWeather.yaml", project=PROJECT)

    # 2) Connect and query
    async with client(project=PROJECT, uuid=USER_ID) as dc:
        responses = await dc.query(query="What is time in India")
        async for r in responses:
            print("Response:", r["content"])

if __name__ == "__main__":
    asyncio.run(main())

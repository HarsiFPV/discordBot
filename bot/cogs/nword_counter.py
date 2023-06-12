"""Cog for n-word counting and storing logic"""
import random
import string
import asyncio
import discord.ext
from typing import List

import discord
from discord.ext import commands

from utils.mongo_instance import database

# Create the n-word lists from ASCII so I don't have to type it.
NWORDS_LIST = [
    (chr(110) + chr(105) + chr(103) + chr(103) + chr(97)),
    (chr(47) + chr(92) + chr(47) + chr(105) + chr(103) + chr(103) + chr(97)),
    (chr(124) + chr(92) + chr(47) + chr(105) + chr(103) + chr(103) + chr(97)),
    (chr(114) + chr(97) + chr(116) + chr(105) + chr(111)),
    (chr(100) + chr(97) + chr(118) + chr(105) + chr(100)),
    (chr(104) + chr(116) + chr(116) + chr(112) + chr(115) + chr(58) + chr(47) + chr(47) + chr(116) + chr(119) + chr(105) + chr(116) + chr(116) + chr(101) + chr(114) + chr(46) + chr(99) + chr(111) + chr(109) + chr(47))
]
HARD_RS_LIST = [
    (chr(110) + chr(105) + chr(103) + chr(103) + chr(101) + chr(114)),
    (chr(47) + chr(92) + chr(47) + chr(105) + chr(103) + chr(103) + chr(101) + chr(114)),
    (chr(124) + chr(92) + chr(47) + chr(105) + chr(103) + chr(103) + chr(101) + chr(114))
]


class NWordCounter(commands.Cog):
    """Commands for n-word count tracking"""

    def __init__(self, bot):
        self.bot = bot

        # Get singleton database connection.
        self.db = database

        self.sacred_n_words = NWordCounter.get_n_words()
        self.sacred_hard_r_words = NWordCounter.get_hard_r_words()
    

    @staticmethod
    def get_n_words() -> List:
        return NWORDS_LIST


    @staticmethod
    def get_hard_r_words() -> List:
        return HARD_RS_LIST
    

    @staticmethod
    def count_nwords(msg: str) -> int:
        """Return occurrences of n-words in a given message"""
        count = 0
        msg = msg.lower().strip().translate(  # Cleanse all whitespaces.
            {ord(char): "" for char in string.whitespace}
        )

        for n_word in NWordCounter.get_n_words():
            count += msg.count(n_word)
        for hard_r in NWordCounter.get_hard_r_words():
            count += msg.count(hard_r)

        return count

    
    def is_black(self, guild_id, author_id) -> bool:
        """Check if user is verified to be black"""
        member = self.db.member_in_database(guild_id, author_id)
        if member["is_black"]:
            return True
        return False
    

    def get_member_nword_count(self, guild_id, member_id) -> int:
        """Return number of n-words said by member if tracked in database"""
        member: object | None = self.db.member_in_database(guild_id, member_id)
        if not member:
            return 0
        return member["nword_count"]
    

    def get_msg_response(self, nword_count: int) -> str:
        """Return bot message response based on n-words said"""
        msg = None
        if nword_count < 5:
            msg = random.choice(
                [
                    "Tu te fous de ma gueule ? :face_with_raised_eyebrow::camera_with_flash:",
                    "APANIAN ? :face_with_raised_eyebrow::camera_with_flash:",
                    "Ratio plus raciste :camera_with_flash:",
                    "Nigga toi même bougnoule :camera_with_flash:",
                    "Je baise ta mère :face_with_raised_eyebrow:",
                    "ESCLAVE DE MERDE"
                ]
            )
        elif nword_count < 25:
            msg = random.choice(
                [
                    "Bro chill with it :camera_with_flash:",
                    "?????????? :camera_with_flash:",
                    "Bro cmon :neutral_face:",
                    "Bro..."
                ]
            )
        elif nword_count < 100:
            msg = random.choice(
                [
                    "Bro wtf :face_with_raised_eyebrow:",
                    ":expressionless:",
                    "CHILL",
                    ":camera_with_flash::camera_with_flash::camera_with_flash:",
                    "..."
                ]
            )
        else:
            msg = random.choice(
                [
                    "I have no words.",
                    "Tf?",
                    ":exploding_head:",
                    ":flushed:",
                    "I'm calling your employer",
                    "This gotta affect your credit score somehow",
                    "Quit this madness",
                    "Bro you ARE the n-word pass",
                    ":farmer:"
                ]
            )
        
        return msg


    #@commands.Cog.listener()
    #async def on_message(self, message):
    #    """Detect n-words"""
    #    if message.author == self.bot.user:  # Ignore reading itself.
    #        return
#
    #    guild = message.guild
    #    msg = message.content
    #    author = message.author  # Should fetch user by ID instead of name.
#
    #    # Ensure guild has its own place in the database.
    #    if not self.db.guild_in_database(guild.id):
    #        self.db.create_database(guild.id, guild.name)
    #
    #    # Bot reaction to any n-word occurrence.
    #    num_nwords = self.count_nwords(msg)
    #    if num_nwords < 5 :
    #        if message.webhook_id:  # Ignore webhooks.
    #            await message.reply("Not a person, I won't count this.")
    #            return
    #        if message.webhook_id:  # Ignore webhooks.
    #            await message.reply("Not a person, I won't count this.")
    #            return
#
    #        if not self.db.member_in_database(guild.id, author.id):
    #            self.db.create_member(guild.id, author.id, author.name)
    #
    #        self.db.increment_nword_count(guild.id, author.id, num_nwords)
#
    #        if self.is_black(guild.id, author.id):  # Don't react to someone already verified.
    #            return
#
    #        # response = f":camera_with_flash: Detected **{num_nwords}** n-words!"
    #        response = self.get_msg_response(nword_count=num_nwords)
    #        await message.author.edit(voice_channel=None)
    #        await message.delete()
    #        await message.reply(f"{message.author.mention} Ciao bozo")
    

    def get_id_from_mention(self, mention: str) -> int:
        """Extract user ID from mention string"""
        # STORED IN DB AS INTEGER, NOT STRING.
        return int(mention.replace("<@!", "").replace("<@", "").replace(">", ""))
    

    def verify_mentions(self, mentions, ctx) -> str:
        """Check if mention being passed into command is valid.
        
        Return error message if not validated
        Return "" if validated
        """
        if not mentions:
            return "Mention not found"
        elif len(mentions) > 1:
            return "Only mention one user"
        
        id_from_mention = mentions[0].id

        # Ensure user is part of guild.
        guild = ctx.guild
        if not guild.get_member(id_from_mention):
            return "User not in server"
        
        return ""
    

    @commands.command()
    async def count(self, ctx, mention=None):
        """Get a person's total n-word count"""
        mentions: list = ctx.message.mentions
        member_id = None
        mention_name = None

        if not mention:  # No mention = get author.
            member_id = ctx.author.id
            mention_name = ctx.author.mention
        else:  # Validate mention.
            invalid_mention_msg = self.verify_mentions(mentions, ctx)
            if invalid_mention_msg:
                await ctx.reply(invalid_mention_msg)
                return
            member_id = mentions[0].id
            mention_name = mention

        # Fetch n-word count of user if they have a count.
        nword_count = self.get_member_nword_count(ctx.guild.id, member_id)
        await ctx.send(f"Total n-word count for {mention_name}: `{nword_count:,}`")
    

    def get_vote_threshold(self, member_count: int) -> int:
        """Return votes required to verify based on server member count"""
        if member_count < 3:
            return 1
        elif member_count < 10:
            return 2
        elif member_count < 50:
            return 3
        elif member_count < 100:
            return 4
        elif member_count < 250:
            return 5
        else:
            return 10
    

    def user_voted_for(self, voter_id: int, member: object) -> bool:
        """Return whether or not user voted for the member"""
        if member is None:
            return False

        voter_list = member["voters"]
        if voter_id in voter_list:
            return True

        return False
    

    @commands.command()
    async def vote(self, ctx, mention=None):
        """Vouch to verify someone's blackness"""
        vote_status_msg = ""
        if not mention:
            vote_status_msg = self.perform_vote(ctx, "vote")
        else:
            vote_status_msg = self.perform_vote(ctx, "vote", mention)
        await ctx.reply(vote_status_msg)


    @commands.command()
    async def unvote(self, ctx, mention=None):
        """Remove a vouch for a person"""
        vote_status_msg = ""
        if not mention:
            vote_status_msg = self.perform_vote(ctx, "unvote")
        else:
            vote_status_msg = self.perform_vote(ctx, "unvote", mention)
        await ctx.reply(vote_status_msg)
    

    def get_vote_return_msgs(self, type, votes, vote_threshold) -> dict:
        """Get vote responses messages based on action"""
        msgs_dict = dict()
        if type == "vote":
            msgs_dict["already_performed_msg"] = f"You already voted for this person!\n({votes}/{vote_threshold}) required votes so far!"
            msgs_dict["error_performed_msg"] = "Couldn't vote for person"
            msgs_dict["success_performed_msg"] = f"Successfully voted!\n({votes}/{vote_threshold}) required votes so far!"
        else:
            msgs_dict["already_performed_msg"] = f"You never voted for this person!\n({votes}/{vote_threshold}) required votes so far!"
            msgs_dict["error_performed_msg"] = "Couldn't unvote person"
            msgs_dict["success_performed_msg"] = f"Successfully removed vote!\n({votes}/{vote_threshold}) required votes so far!"
        return msgs_dict



    def perform_vote(self, ctx, type, mention=None) -> str:
        """Main logic for voting and unvoting"""
        mentions: list = ctx.message.mentions
        mention_id = None
        mention_name = None

        # Must mention someone.
        invalid_mention_msg = self.verify_mentions(mentions, ctx)
        if invalid_mention_msg:
            return invalid_mention_msg
        
        mention_id = mentions[0].id
        mention_name = mentions[0].name
        
        # Can't vote for yourself.
        if mention_id == ctx.author.id:
            return "You can't vote/unvote for yourself bozo"

        # Create member if not already in database.
        member = self.db.member_in_database(ctx.guild.id, mention_id)
        if not member:
            self.db.create_member(ctx.guild.id, mention_id, mention_name)
        
        member_count = len(ctx.guild.members)
        vote_threshold = self.get_vote_threshold(member_count)
        votes = len(member["voters"])

        # Check whether user has a vote casted on them already.
        if type == "vote":
            if self.user_voted_for(ctx.author.id, member):
                msg = self.get_vote_return_msgs(type, votes, vote_threshold)
                return msg["already_performed_msg"]
        else:
            if not self.user_voted_for(ctx.author.id, member):
                msg = self.get_vote_return_msgs(type, votes, vote_threshold)
                return msg["already_performed_msg"]
            
        # Perform and let know result.
        voted = self.db.cast_vote(type, ctx.guild.id, vote_threshold, ctx.author.id, mention_id)
        member = self.db.member_in_database(ctx.guild.id, mention_id)
        votes = len(member["voters"])
        if not voted:
            msg = self.get_vote_return_msgs(type, votes, vote_threshold)
            return msg["error_performed_msg"]
        else:
            msg = self.get_vote_return_msgs(type, votes, vote_threshold)
            return msg["success_performed_msg"]


    @commands.command()
    async def whoblack(self, ctx):
        """See who's verified black in this server"""
        member_list = [
            member["name"]
            for member in self.db.get_member_list(ctx.guild.id)
            if member["is_black"]
        ]
        msg = "Verified black members in this server:\n"
        if len(member_list) == 0:
            msg += "`None`"
        else:
            for member in member_list:
                msg += f" `{member}`"
        await ctx.send(msg)
    

    @commands.command()
    async def whohaspass(self, ctx):
        """See who has an n-word pass in this server"""
        member_list = [
            member["name"]
            for member in self.db.get_member_list(ctx.guild.id)
            if member["has_pass"]
        ]
        msg = "Verified pass holders in this server:\n"
        if len(member_list) == 0:
            msg += "`None`"
        else:
            for member in member_list:
                msg += f" `{member}`"
        await ctx.send(msg)
    

    @commands.command()
    async def passes(self, ctx, mention=None):
        """See your or someone else's total passes available"""
        mentions: list = ctx.message.mentions

        if not mention:  # Passes for author.
            member = self.db.member_in_database(ctx.guild.id, ctx.author.id)
            if not member:
                self.db.create_member(ctx.guild.id, ctx.author.id, ctx.author.name)
                member = self.db.member_in_database(ctx.guild.id, ctx.author.id)
            await ctx.send(f"Your n-word passes: `{member['passes']}`")
        else:  # Passes for mentioned member.
            invalid_mention_msg = self.verify_mentions(mentions, ctx)
            if invalid_mention_msg:
                await ctx.send(invalid_mention_msg)
                return

            mention_id = mentions[0].id
            mention_name = mentions[0].name
            member = self.db.member_in_database(ctx.guild.id, mention_id)
            if not member:
                self.db.create_member(ctx.guild.id, mention_id, mention_name)
                member = self.db.member_in_database(ctx.guild.id, mention_id)
            await ctx.send(f"N-word passes for {mention_name}: `{member['passes']}`")


    @commands.command()
    async def givepass(self, ctx):
        """Give n-word passes to another person"""
        await ctx.send("Feature coming soon!")

async def setup(bot):
    await bot.add_cog(NWordCounter(bot))


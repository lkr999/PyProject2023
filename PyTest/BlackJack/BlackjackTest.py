import random

class Card:
    def __init__(self, suit, value):
        self.suit = suit
        self.value = value

    def __repr__(self):
        return f"{self.value} of {self.suit}"

class Deck:
    def __init__(self):
        self.cards = [Card(s, v) for s in ["Spades", "Clubs", "Hearts", "Diamonds"] for v in range(1, 14)]

    def shuffle(self):
        if len(self.cards) > 1:
            random.shuffle(self.cards)

    def deal(self):
        if len(self.cards) > 1:
            return self.cards.pop(0)

class Hand:
    def __init__(self):
        self.cards = []

    def add_card(self, card):
        self.cards.append(card)

    def get_value(self):
        value = 0
        has_ace = False
        for card in self.cards:
            if card.value > 10:
                value += 10
            elif card.value == 1:
                has_ace = True
                value += 11
            else:
                value += card.value

        if has_ace and value > 21:
            value -= 10

        return value

    def __repr__(self):
        return str(self.cards)

class Blackjack:
    def __init__(self):
        self.deck = Deck()
        self.deck.shuffle()
        self.player_hand = Hand()
        self.dealer_hand = Hand()

    def deal_initial_cards(self):
        self.player_hand.add_card(self.deck.deal())
        self.dealer_hand.add_card(self.deck.deal())
        self.player_hand.add_card(self.deck.deal())
        self.dealer_hand.add_card(self.deck.deal())

    def player_turn(self):
        while self.player_hand.get_value() < 21:
            action = input("Do you want to hit or stand? ")
            if action == "hit":
                self.player_hand.add_card(self.deck.deal())
                print(f"Your hand: {self.player_hand}")
            elif action == "stand":
                break
            else:
                print("Invalid action. Please try again.")
                continue

    def dealer_turn(self):
        while self.dealer_hand.get_value() < 17:
            self.dealer_hand.add_card(self.deck.deal())
            print(f"Dealer's hand: {self.dealer_hand}")

    def determine_winner(self):
        player_value = self.player_hand.get_value()
        dealer_value = self.dealer_hand.get_value()

        if player_value > 21:
            return "Dealer wins! You went over 21."
        elif dealer_value > 21:
            return "You win! Dealer went over 21."
        elif player_value > dealer_value:
            return "You win!"
        elif dealer_value > player_value:
            return "Dealer wins!"
        else:
            return "It's a tie!"

    def play(self):
        self.deal_initial_cards()
        print(f"Your hand: {self.player_hand}")
        print(f"Dealer's up card: {self.dealer_hand.cards[1]}")
        self.player_turn()
        self.dealer_turn()
        print(self.determine_winner())

game = Blackjack()
game.play()

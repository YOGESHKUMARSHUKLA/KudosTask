Feature: Create and View Post on Kudos

  Scenario Outline: Login with Alphabetical,Numerical,SpecialChars - username and Password

    Given the user is on Kudos Login Page
    When user login with "<username>" username and "<password>" password
    Then user view the post
    And user like the post
    When user create post


    Examples:

      | TestName | username | password     |
      | Test1    | Test     | TestPassword |
      | Test2    | 123456   | 123456       |
      | Test3    | #$#%%    | #$#$$$#$#$&* |

  Scenario Outline: Login with Alphabetical - username and Password and View Post

    Given the user is on Kudos Login Page
    When user login with "<username>" username and "<password>" password
    Then user view the post
    Examples:

      | TestName | username | password     |
      | Test1    | Test     | TestPassword |

  Scenario Outline: Login with Alphabetical - username and Password and Like the post

    Given the user is on Kudos Login Page
    When user login with "<username>" username and "<password>" password
    Then user like the post
    Examples:

      | TestName | username | password     |
      | Test1    | Test     | TestPassword |

  Scenario Outline: Login with Alphabetical - username and Password and create Post

    Given the user is on Kudos Login Page
    When user login with "<username>" username and "<password>" password
    Then user create post
    Examples:

      | TestName | username | password     |
      | Test1    | Test     | TestPassword |

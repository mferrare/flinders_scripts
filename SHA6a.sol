pragma solidity >=0.4.22 <0.6.0;
/*
 * Implementation of ABC Pty Ltd Shareholders Agreement as a smart contract
 * Uses ERC-20 tokens as shares
 */
// This part is required to implement the ERC-20 standard
contract Token {
    function totalSupply() public view returns (uint256 supply) {}
    function balanceOf(address _owner) public view returns (uint256 balance) {}
    function transfer(address _to, uint256 _value) public returns (bool success) {}
    function transferFrom(address _from, address _to, uint256 _value) public returns (bool success) {}
    function approve(address _spender, uint256 _value) public returns (bool success) {}
    function allowance(address _owner, address _spender) public view returns (uint256 remaining) {}
    event Transfer(address indexed _from, address indexed _to, uint256 _value);
    event Approval(address indexed _owner, address indexed _spender, uint256 _value);   
}
/*
 * This Shareholders Agreement proper is defined in this smart contract
 */
contract ABC_ShareholderAgreement is Token {
    address public sole_director;
    address public Mark_Joseph_Ferraretto = 0x33f2e15d377a3B34C22CE2d74Aa81C7A44bCd1Ba;
    mapping (address => uint) public share_registry;
    mapping (address => mapping (address => uint256)) allowed;
    string public status_message;
    uint256 private total_shares;
    
    // A 'constructor' is executed automatically as the smart contract is deployed to
    // the blockchain.  So, it's a good place to set initial values.
    // We use the constructor to implement Clause 1(a) of the Agreement and also to tell
    // the smart contract that zero shares have been issued.
    constructor() public {
        // Clause 1(a) of the Agreement
        sole_director = Mark_Joseph_Ferraretto;
        total_shares = 0;
    }
    
    // This function implements Clause 2 of the Agreement
    function issueShares(address receiver, uint amount) public {
        if (msg.sender != sole_director) {
            status_message = "Not the sole director";
            return;
        } else {
            share_registry[receiver] += amount;
            total_shares += amount;
            status_message = "Shares were issued";
        }
    }
    // This function implements Clause 3 of the Agreement
    function transfer( address _to, uint256 _value ) public returns (bool success) {
        if (share_registry[msg.sender] >= _value && _value > 0) {
            share_registry[msg.sender] -= _value;
            share_registry[_to] += _value;
            emit Transfer(msg.sender, _to, _value);
            status_message = "Shares transferred";
            return true;
        } else { 
            status_message = "Shares not transferred";
            return false;
        }
    }
    /*
     * The following functions are used to implement the ERC-20
     * protocol.  This enables shareholders to transfer shares from
     * their wallets instead of having to interact with the smart
     * contract directly
     */
    string public name = "ABC Pty Ltd Ordinary Shares";
    string public symbol = "ABC_ORD";
    uint8 public decimals = 0;
    
    function totalSupply() public view returns (uint256) {
        return total_shares;
    }
    
    function balanceOf(address _member) public view returns (uint256 balance) {
        return share_registry[_member];
    }
    
    function transferFrom(address _from, address _to, uint256 _value) public returns (bool success) {
        if (share_registry[_from] >= _value && allowed[_from][msg.sender] >= _value && _value > 0) {
            share_registry[_to] += _value;
            share_registry[_from] -= _value;
            allowed[_from][msg.sender] -= _value;
            emit Transfer(_from, _to, _value);
            return true;
        } else {
            return false;
        }
    }
    
    function approve(address _spender, uint256 _value) public returns (bool success) {
        allowed[msg.sender][_spender] = _value;
        emit Approval(msg.sender, _spender, _value);
        return true;
    }
    function allowance(address _owner, address _spender) public view returns (uint256 remaining) {
      return allowed[_owner][_spender];
    }
    
    event Transfer(address indexed _from, address indexed _to, uint256 _value);
    event Approval(address indexed _owner, address indexed _spender, uint256 _value);
}
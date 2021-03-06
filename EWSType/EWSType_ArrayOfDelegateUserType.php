<?php namespace Exchange\EWSType; use Exchange\EWSType;
/**
 * Contains EWSType_ArrayOfDelegateUserType.
 */

/**
 * Contains the identities of delegates to add to or update in a mailbox.
 *
 * @package php-ews\Types
 */
class EWSType_ArrayOfDelegateUserType extends EWSType
{
    /**
     * Identifies a single delegate to add to or update in a mailbox.
     *
     * @since Exchange 2007 SP1
     *
     * @var EWSType_DelegateUserType
     */
    public $DelegateUser;
}

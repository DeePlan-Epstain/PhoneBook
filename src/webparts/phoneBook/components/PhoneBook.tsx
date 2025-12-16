import * as React from "react";
import styles from "./PhoneBook.module.scss";
import { jss } from "./models/jss";
import { theme } from "./models/theme";
import { IPhoneBookProps } from "./IPhoneBookProps";
import { IPhoneBookState } from "./IPhoneBookState";
import { ThemeProvider, StylesProvider } from "@material-ui/core/styles";
import { TextField, withStyles } from "@material-ui/core";
import Autocomplete from "@material-ui/lab/Autocomplete";
import SearchIcon from "@material-ui/icons/Search";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { Pagination } from "@material-ui/lab";
import "./workbench.css";
import Skeleton from "@material-ui/lab/Skeleton";
import AddIcon from "@material-ui/icons/Add";
import IconButton from "@material-ui/core/IconButton";
import Dialog from "@material-ui/core/Dialog";
import { Web } from "@pnp/sp/webs";
import { Item } from "@pnp/sp/items";
import Button from "@material-ui/core/Button";

interface IContact {
    displayName: string;
    firstName: string;
    lastName: string;
    phoneNumber: string;
    Email: string;
    ID: string;
    Division: string;
    Role: string;
    secondPhoneNumber: string;
}

const WithStylesTextField = withStyles({
    root: {
        color: "white",
        background: "inherit",
        height: "50px",
        "& .MuiInputLabel-outlined": {},
        "& .MuiAutocomplete-popupIndicatorOpen": {
            transform: "rotate(0)",
        },
        "& .MuiAutocomplete-listBox": {
            direction: "rtl !important",
        },
        "& label.Mui-focused": {
            color: "white",
        },
        "& .MuiInput-underline:after": {
            borderBottomColor: "#939393",
        },
        "& .MuiOutlinedInput-root": {
            height: "37px",
            backgroundColor: "white",
            padding: "0px 4px",
            borderRadius: "50px",
            boxShadow: "1px 1px 2px 0px #0000001A",
            "& fieldset": {
                border: "1px solid #D3D3D380",
            },
            "&:hover fieldset": {
                borderColor: "#939393",
            },
            "&.Mui-focused fieldset": {
                borderColor: "#939393",
            },
            "& .MuiIconButton-root": {
                padding: "4px",
            },
        },
    },
})(TextField);

export default class PhoneBook extends React.Component<IPhoneBookProps, IPhoneBookState> {
    private sp = spfi().using(SPFx(this.props.context));

    constructor(props: IPhoneBookProps) {
        super(props);
        this.state = {
            IsAutocompleteOptionsOpen: false,
            SearchValue: " ",
            ContactsOptions: [],
            DisplayOptions: [],
            IsShowNewRowButton: false,
            CurrentPage: 1,
            NumOfPages: 0,
            PageSize: 8,
            CurrPageDisplayOptions: [],
            Contacts: [],
            PageNumber: 1,
            isLoading: true,
            errorMsgDisplay: false,
            imgSkeletonFlag: true,
            isModalOpen: false,
            newUser: {
                displayName: "",
                firstName: "",
                lastName: "",
                phoneNumber: "",
                Email: "",
            },
            errors: {},
            titleAndUrl: [],
        };
    }

    componentDidMount(): void {
        this.getEmployeeData();
        let width = window.innerWidth;

        if (width < 1020) {
            this.setState({ ...this.state, PageSize: 3 });
        }
    }

    toggleModal = () => {
        this.setState({ isModalOpen: !this.state.isModalOpen });
    };

    addNewUser = async () => {
        const { newUser } = this.state;
        const errors: {
            firstName?: string;
            lastName?: string;
            Email?: string;
            phoneNumber?: string;
        } = {};

        // Validate each field and set error messages
        if (!newUser.firstName) errors.firstName = "שם פרטי שדה חובה";
        if (!newUser.lastName) errors.lastName = "שם משפחה שדה חובה";
        if (!newUser.Email) errors.Email = "אימייל שדה חובה";
        if (!newUser.phoneNumber) errors.phoneNumber = "טלפון שדה חובה";
        // Check if there are errors
        if (Object.keys(errors).length > 0) {
            // Set the errors in state and stop execution if errors exist
            this.setState({ errors });
            return;
        }

        // No errors, proceed with adding the user
        try {
            const newItem = await this.sp.web.lists.getById(this.props.PhoneBookTableId).items.add({
                displayName: `${newUser.firstName} ${newUser.lastName}`,
                firstName: newUser.firstName,
                lastName: newUser.lastName,
                phoneNumber: newUser.phoneNumber,
                Email: newUser.Email,
            });

            const addedUser = {
                ID: newItem.data.ID,
                displayName: newUser.displayName,
                firstName: newUser.firstName,
                lastName: newUser.lastName,
                phoneNumber: newUser.phoneNumber,
                Email: newUser.Email,
            };

            // Update state with the newly added user and reset form
            this.setState({
                Contacts: [...this.state.Contacts, addedUser],
                newUser: {
                    displayName: "",
                    firstName: "",
                    lastName: "",
                    phoneNumber: "",
                    Email: "",
                },
                errors: {}, // Clear errors after submission
                isModalOpen: false,
            });

            // Reload or fetch updated data if needed
            this.getEmployeeData();
            console.log("User added successfully:", addedUser);
        } catch (error) {
            console.error("Error adding new user:", error);
        }
    };

    getEmployeeData = async () => {
        try {
            const listItems = await this.sp.web.lists.getById(this.props.PhoneBookTableId).items.top(1000)();

            const mappedItems: IContact[] = listItems
                .filter((item: any) => item.show === true)
                .sort((a, b) => {
                    const nameA = a.firstName.toUpperCase(); // ignore upper and lowercase
                    const nameB = b.firstName.toUpperCase(); // ignore upper and lowercase
                    if (nameA < nameB) {
                        return -1;
                    }
                    if (nameA > nameB) {
                        return 1;
                    }

                    // names must be equal
                    return 0;
                })
                .map((item: any) => ({
                    displayName: item.displayName || `${item.firstName} ${item.lastName}`,
                    firstName: item.firstName,
                    lastName: item.lastName,
                    phoneNumber: item.phoneNumber,
                    Email: item.Email,
                    ID: item.ID,
                    Division: item.Division !== null ? item.Division : "",
                    Role: item.Role !== null ? item.Role : "",
                    secondPhoneNumber: item.secondPhoneNumber !== null ? item.secondPhoneNumber : "",
                }));
            let s: { Title: any; url: any }[] = [];

            this.setState(
                {
                    titleAndUrl: s,
                    Contacts: mappedItems,
                    ContactsOptions: mappedItems,
                },
                () => {
                    this.FilterDisplayedOptions();
                    setTimeout(() => {
                        this.setState({
                            isLoading: false,
                        });
                    }, 1500);
                }
            );
        } catch (err) {
            console.error("Error fetching employee data:", err);
            this.setState({
                errorMsgDisplay: true,
            });
        }
    };

    handleNewValue = (event: any) => {
        this.setState({
            SearchValue: event.target.value ? event.target.value : "",
            IsAutocompleteOptionsOpen: event.target.value ? true : false,
        });
    };

    FilterDisplayedOptions = () => {
        const { SearchValue, PageSize, Contacts } = this.state;
        let FilteredContacts: Array<IContact> = Contacts;

        // Convert the search value to lowercase for case-insensitive comparison
        const lowerCaseSearchValue = SearchValue.trim().toLowerCase();

        // Filter by search value across multiple fields

        if (lowerCaseSearchValue !== "") {
            FilteredContacts = FilteredContacts.filter((contact: IContact) => {
                const fieldsToSearch = [
                    contact.phoneNumber,
                    contact.Email,
                    contact.displayName,
                    contact.lastName,
                    contact.firstName,
                    contact.Division,
                    contact.Role,
                    contact.secondPhoneNumber,
                ];

                // Check if the search value exists in any of the fields
                return fieldsToSearch.some((field) => (field ? field.toLowerCase().includes(lowerCaseSearchValue) : false));
            });
        }

        // Pagination logic
        const indexOfLastOption = 1 * PageSize;
        const indexOfFirstOption = indexOfLastOption - PageSize;
        const CurrPageDisplayOptions = FilteredContacts.slice(indexOfFirstOption, indexOfLastOption);
        const NumOfPages = Math.ceil(FilteredContacts.length / PageSize);

        this.setState({
            DisplayOptions: FilteredContacts,
            IsAutocompleteOptionsOpen: false,
            IsShowNewRowButton: FilteredContacts.length ? false : true,
            NumOfPages,
            CurrPageDisplayOptions,
            CurrentPage: 1,
        });
    };

    OnSetCurrentPage = async (event: object, page: number) => {
        const { PageSize, DisplayOptions } = this.state;
        const indexOfLastOption = page * PageSize;
        const indexOfFirstOption = indexOfLastOption - PageSize;
        const CurrPageDisplayOptions = DisplayOptions.slice(indexOfFirstOption, indexOfLastOption);

        this.setState({
            CurrentPage: page,
            CurrPageDisplayOptions,
            PageNumber: page,
        });
    };

    public render(): React.ReactElement<IPhoneBookProps> {
        const StyledPagination = withStyles((theme) => ({
            ul: {
                "& .MuiPaginationItem-root": {
                    color: "#1369ce",
                },
                "& .MuiPaginationItem-textPrimary.Mui-selected": {
                    backgroundColor: "1369ce",
                    color: "#fff",
                },
            },
        }))(Pagination);

        return (
            <section className={` ${styles.PhoneBookMainCon}`}>
                <div id="phonebook" className={`${styles.PhoneBookBackGroundImage}`}>
                    <div className={`${styles.PhoneBookRightContainer}`}>
                        <div className={`${styles.PhoneBookTitlesContainer}`}>
                            <span className={styles.PhoneBookNewsTitle}>ספר טלפונים</span>
                            <img className={`${styles.PhoneBookLine20}`} src={require("../assets/Line20.svg")} alt="" />
                        </div>
                    </div>
                    <div className={`${styles.phoneBookLeftContainer}`}>
                        <div className={styles.PhoneBookContainer}>
                            <div className={styles.PhoneBook} dir="rtl">
                                <div className={styles.Search} dir="rtl">
                                    <StylesProvider jss={jss}>
                                        <ThemeProvider theme={theme}>
                                            <div className="search-filter-wrapper">
                                                <Autocomplete
                                                    clearOnBlur={false}
                                                    style={{ direction: "rtl", textAlign: "center" }}
                                                    onClose={(event, reason) => {
                                                        this.setState({ IsAutocompleteOptionsOpen: false });
                                                    }}
                                                    open={this.state.IsAutocompleteOptionsOpen}
                                                    noOptionsText={"אין תוצאות"}
                                                    onInputChange={(event, newInputValue) => {
                                                        // Update search value on every input change
                                                        this.setState(
                                                            { SearchValue: newInputValue },
                                                            this.FilterDisplayedOptions
                                                        );
                                                    }}
                                                    onChange={(event, newValue) => {
                                                        if (newValue) {
                                                            this.setState(
                                                                {
                                                                    SearchValue: newValue.displayName, // Use displayName or another key as needed
                                                                    IsAutocompleteOptionsOpen: false,
                                                                },
                                                                this.FilterDisplayedOptions
                                                            );
                                                        }
                                                    }}
                                                    popupIcon={<SearchIcon style={{ color: "white[500]" }} />}
                                                    options={this.state.ContactsOptions}
                                                    getOptionLabel={(option) => option.displayName || ""}
                                                    filterOptions={(options, { inputValue }) => {
                                                        const lowerCaseInputValue = inputValue.trim().toLowerCase();
                                                        return options.filter((option) => {
                                                            const fieldsToSearch = [
                                                                option.phoneNumber,
                                                                option.Email,
                                                                option.displayName,
                                                                option.lastName,
                                                                option.firstName,
                                                            ];

                                                            return fieldsToSearch.some((field) =>
                                                                field ? field.toLowerCase().includes(lowerCaseInputValue) : false
                                                            );
                                                        });
                                                    }}
                                                    renderInput={(params) => (
                                                        <WithStylesTextField
                                                            style={{ direction: "rtl", float: "right" }}
                                                            {...params}
                                                            placeholder="חיפוש"
                                                            variant="outlined"
                                                            InputLabelProps={{
                                                                style: { color: "#fff", textAlign: "center" },
                                                            }}
                                                        />
                                                    )}
                                                    size="medium"
                                                    fullWidth
                                                />
                                            </div>
                                        </ThemeProvider>
                                    </StylesProvider>
                                    {/* <div>
                    <IconButton aria-label="add" onClick={this.toggleModal}>
                      <AddIcon fontSize="large" />
                    </IconButton>
                  </div> */}
                                </div>
                                <div className={styles.ItemsContainer}>
                                    {this.state.CurrPageDisplayOptions.map((item, index) => {
                                        return !this.state.isLoading ? (
                                            <div className={styles.UserItem} key={item.ID}>
                                                <div className="our-team">
                                                    <div className="picture">
                                                        <img
                                                            className="img-fluid"
                                                            src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=${item?.Email}`}
                                                        />
                                                    </div>

                                                    <div className="team-content">
                                                        {/* <h3 className="name">{item.displayName}</h3>
                            <h3 className="name">{item.phoneNumber}</h3>
                            {item.Division && <h3 className="name">{item.Division}</h3>}
                            {item.Role && <h3 className="name">{item.Role}</h3>}
                            {item.secondPhoneNumber && <h3 className="name">{item.secondPhoneNumber}</h3>} */}

                                                        {/* <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: '0.35rem', width: '100%', height: '100%' }}> */}
                                                        {/* FontSize 14px to all of them, image 80x80px */}
                                                        <span className="textInfo" style={{ fontWeight: "bold" }}>
                                                            {item.displayName}
                                                        </span>
                                                        <span className="textInfo">{item.phoneNumber}</span>
                                                        {item.secondPhoneNumber && (
                                                            <span className="textInfo">שלוחה: {item.secondPhoneNumber}</span>
                                                        )}
                                                        {item.Division && <span className="textInfo">{item.Division}</span>}
                                                        {item.Role && <span className="textInfo">{item.Role}</span>}
                                                        {/* </span> */}
                                                    </div>
                                                    <ul className="social">
                                                        <li>
                                                            <a href={`mailto:${item.Email}`}>שלח מייל</a>
                                                        </li>
                                                    </ul>
                                                </div>
                                            </div>
                                        ) : (
                                            <div className={styles.Item} key={index}>
                                                <Skeleton className={styles.ItemImg} variant="circle" width={40} height={40} />
                                                <Skeleton className={styles.ItemTitle} variant="text" width={210} height={20} />
                                                <Skeleton className={styles.ItemLine} variant="text" width={210} height={20} />
                                            </div>
                                        );
                                    })}
                                </div>
                                {this.state.DisplayOptions.length ? (
                                    <div className={styles.PaginationContainer}>
                                        <ThemeProvider theme={theme}>
                                            <StyledPagination
                                                onChange={this.OnSetCurrentPage}
                                                page={this.state.CurrentPage}
                                                count={this.state.NumOfPages}
                                                skin-primary-color="secondary"
                                                disabled={this.state.DisplayOptions.length === 0}
                                                classes={{ ul: styles.pagination }}
                                            />
                                        </ThemeProvider>
                                    </div>
                                ) : null}
                                {!this.state.DisplayOptions.length && !this.state.isLoading ? (
                                    <h3 style={{ textAlign: "center" }}>אין תוצאות</h3>
                                ) : null}
                            </div>
                        </div>
                    </div>
                </div>
                {/* <Dialog
          open={this.state.isModalOpen}
          onClose={this.toggleModal}
          aria-labelledby="form-dialog-title"
          maxWidth="lg"
        >
          <div className="container">
            <div className="text">הוספת לקוח חדש</div>
            <form>
              <div className="form-row">
                <div className="input-data">
                  <input
                    type="text"
                    required
                    value={this.state.newUser.firstName}
                    onChange={(e) =>
                      this.setState({
                        newUser: {
                          ...this.state.newUser,
                          firstName: e.target.value,
                        },
                        errors: {
                          ...this.state.errors,
                          firstName: "",
                        },
                      })
                    }
                  />
                  <div className="underline"></div>
                  <label>שם פרטי</label>
                  {this.state.errors.firstName && (
                    <span className="error-message">
                      {this.state.errors.firstName}
                    </span>
                  )}
                </div>
                <div className="input-data">
                  <input
                    type="text"
                    required
                    value={this.state.newUser.lastName}
                    onChange={(e) =>
                      this.setState({
                        newUser: {
                          ...this.state.newUser,
                          lastName: e.target.value,
                        },
                        errors: {
                          ...this.state.errors,
                          lastName: "",
                        },
                      })
                    }
                  />
                  <div className="underline"></div>
                  <label>שם משפחה</label>
                  {this.state.errors.lastName && (
                    <span className="error-message">
                      {this.state.errors.lastName}
                    </span>
                  )}
                </div>
              </div>

              <div className="form-row">
                <div className="input-data">
                  <input
                    type="text"
                    required
                    value={this.state.newUser.Email}
                    onChange={(e) =>
                      this.setState({
                        newUser: {
                          ...this.state.newUser,
                          Email: e.target.value,
                        },
                        errors: {
                          ...this.state.errors,
                          Email: "",
                        },
                      })
                    }
                  />
                  <div className="underline"></div>
                  <label>אימייל</label>
                  {this.state.errors.Email && (
                    <span className="error-message">
                      {this.state.errors.Email}
                    </span>
                  )}
                </div>

                <div className="input-data">
                  <input
                    type="text"
                    required
                    value={this.state.newUser.phoneNumber}
                    onChange={(e) =>
                      this.setState({
                        newUser: {
                          ...this.state.newUser,
                          phoneNumber: e.target.value,
                        },
                        errors: {
                          ...this.state.errors,
                          phoneNumber: "",
                        },
                      })
                    }
                  />
                  <div className="underline"></div>
                  <label>טלפון</label>
                  {this.state.errors.phoneNumber && (
                    <span className="error-message">
                      {this.state.errors.phoneNumber}
                    </span>
                  )}
                </div>
              </div>

              <div className="form-row" style={{ width: "96%", margin: "auto" }}>
                <StylesProvider jss={jss}>
                  <ThemeProvider theme={theme}>
                    <Autocomplete
                      id="country-select-demo"
                      onChange={(event, newValue) => {
                        this.setState({
                          companyName: newValue
                        });

                      }}
                      value={this.state.companyName}
                      options={this.state.companyNames}

                      renderInput={(params) => (

                        <TextField  {...params} label="בחירת חברה" />
                      )}
                      fullWidth
                    />
                  </ThemeProvider>
                </StylesProvider>
              </div>
              <div className="form-row">
                <div>
                  <Button
                    style={buttonStyle}
                    className="submit-btn"
                    onClick={this.addNewUser}
                  >
                    שמירה
                  </Button>
                </div>
                <div>
                  <Button
                    style={buttonStyle}
                    className="cancel-btn"
                    onClick={this.toggleModal}
                  >
                    ביטול
                  </Button>
                </div>
              </div>
            </form>
          </div>
        </Dialog> */}
            </section>
        );
    }
}
